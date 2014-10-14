package Email::Outlook::Message;
=head1 NAME

Email::Outlook::Message.pm - Read Outlook .msg files

=head1 SYNOPSIS

  use Email::Outlook::Message;

  my $msg = new Email::Outlook::Message $filename, $verbose;
  my $mime = $msg->to_email_mime;
  $mime->as_string;

=head1 DESCRIPTION

Parses .msg message files as produced by Microsoft Outlook.

=head1 METHODS

=over 8

=item B<new($msg, $verbose)>

Parse the file pointed at by $msg. Set $verbose to a true value to
print information about skipped parts of the .msg file on STDERR.

=item B<to_email_mime>

Output result as an Email::MIME object.

=back

=head1 BUGS

Not all data that is in the .msg file is converted. There are some
parts whose meaning escapes me, although more documentation on MIME
properties is available these days. Other parts do not make sense outside
of Outlook and Exchange.

GPG signed mail is not processed correctly. Neither are attachments of
type 'appledoublefile'.

It would be nice if we could write .MSG files too, but that will require
quite a big rewrite.

=head1 AUTHOR

Matijs van Zuijlen, C<matijs@matijs.net>

=head1 COPYRIGHT AND LICENSE

Copyright 2002--2014 by Matijs van Zuijlen

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
use strict;
use warnings;
use 5.006;
use vars qw($VERSION);
$VERSION = "0.917";

use Email::Simple;
use Email::MIME::Creator;
use Email::Outlook::Message::AddressInfo;
use Email::Outlook::Message::Attachment;
use Carp;
use base 'Email::Outlook::Message::Base';

our $skipheaders = {
  map { uc($_) => 1 }
  "MIME-Version",
  "Content-Type",
  "Content-Transfer-Encoding",
  "X-Mailer",
  "X-Msgconvert",
  "X-MS-Tnef-Correlator",
  "X-MS-Has-Attach"
};

our $MAP_SUBITEM_FILE = {
  '1000' => "BODY_PLAIN",      # Body
  '1009' => "BODY_RTF",        # Compressed-RTF version of body
  '1013' => "BODY_HTML",       # HTML Version of body
  '0037' => "SUBJECT",         # Subject
  '0047' => "SUBMISSION_ID",   # Seems to contain the date
  '007D' => "HEAD",            # Full headers
  '0C1A' => "FROM",            # From: Name
  '0C1E' => "FROM_ADDR_TYPE",  # From: Address type
  '0C1F' => "FROM_ADDR",       # From: Address
  '0E04' => "TO",              # To: Names
  '0E03' => "CC",              # Cc: Names
  '1035' => "MESSAGEID",       # Message-Id
  '1039' => "REFERENCES",      # References: Header
  '1042' => "INREPLYTO",       # In reply to Message-Id
  '3007' => 'DATE2ND',         # Creation Time
  '0039' => 'DATE1ST',         # Outlook sent date
};

#
# Main body of module
#

sub new {
  my $class = shift;
  my $file = shift or croak "File name is required parameter";
  my $verbose = shift;

  my $self = $class->_empty_new;

  $self->{EMBEDDED} = 0;

  my $msg = OLE::Storage_Lite->new($file);
  my $pps = $msg->getPpsTree(1);
  $pps or croak "Parsing $file as OLE file failed";
  $self->_set_verbosity($verbose);
  # TODO: Use separate object as parser?
  $self->_process_pps($pps);

  return $self;
}

sub _empty_new {
  my $class = shift;

  return bless {
    ADDRESSES => [], ATTACHMENTS => [], FROM_ADDR_TYPE => "",
    HAS_UNICODE => 0, VERBOSE => 0, EMBEDDED => 1
  }, $class;
}

sub to_email_mime {
  my $self = shift;

  my $bodymime;
  my $mime;

  my @parts;

  if ($self->{BODY_PLAIN}) { push(@parts, $self->_create_mime_plain_body()); }
  if ($self->{BODY_HTML}) { push(@parts, $self->_create_mime_html_body()); }
  if ($self->{BODY_RTF}) { push(@parts, $self->_create_mime_rtf_body()); }

  if ((scalar @parts) > 1) {
    for (@parts) { $self->_clean_part_header($_) };

    $bodymime = Email::MIME->create(
      attributes => {
        content_type => "multipart/alternative",
        encoding => "8bit",
      },
      parts => \@parts
    );
  } elsif ((@parts) == 1) {
    $bodymime = $parts[0];
  } else {
    $bodymime = $self->_create_mime_plain_body();
  }

  if (@{$self->{ATTACHMENTS}}>0) {
    $self->_clean_part_header($bodymime);
    my $mult = Email::MIME->create(
      attributes => {
        content_type => "multipart/mixed",
        encoding => "8bit",
      },
      parts => [$bodymime],
    );
    foreach my $att (@{$self->{ATTACHMENTS}}) {
      $self->_SaveAttachment($mult, $att);
    }
    $mime = $mult;
  } else {
    $mime = $bodymime;
  }

  #$mime->header_set('Date', undef);
  $self->_SetHeaderFields($mime);
  $self->_copy_header_data($mime);

  return $mime;
}

#
# Below are functions that walk the PPS tree. This is simply a tree walk.
# It's not really recursive (except when an attachment contains a .msg
# file), since the tree is shallow (max. 1 subdirectory deep).
#
# The structure is as follows:
#
# Root
#   Items with properties of the e-mail
#   Dirs containting adresses
#     Items with properties of the address
#   Dirs containing Attachments
#     Items with properties of the attachment (including its data)
#     Dir that is itself a .msg file (if the attachment is an email).
#

sub _property_map {
  return $MAP_SUBITEM_FILE;
}

#
# Process a subdirectory. This is either an address or an attachment.
#
sub _process_subdirectory {
  my ($self, $pps) = @_;

  $self->_extract_ole_date($pps);

  my $name = $self->_get_pps_name($pps);

  if ($name =~ '__recip_version1 0_ ') { # Address of one recipient
    $self->_process_address($pps);
  } elsif ($name =~ '__attach_version1 0_ ') { # Attachment
    $self->_process_attachment($pps);
  } else {
    $self->_warn_about_unknown_directory($pps);
  }
  return;
}

#
# Process a subdirectory that contains an email address.
#
sub _process_address {
  my ($self, $pps) = @_;

  my $addr_info = Email::Outlook::Message::AddressInfo->new($pps,
    $self->{VERBOSE});

  push @{$self->{ADDRESSES}}, $addr_info;
  return;
}

#
# Process a subdirectory that contains an attachment.
#
sub _process_attachment {
  my ($self, $pps) = @_;

  my $attachment = Email::Outlook::Message::Attachment->new($pps,
    $self->{VERBOSE});
  push @{$self->{ATTACHMENTS}}, $attachment;
  return;
}

#
# Header length of the property stream depends on whether the Message
# object is embedded or not.
#
sub _property_stream_header_length {
  my $self = shift;
  return ($self->{EMBEDDED} ?  24 : 32)
}

#
# Helper functions
#

#
# Extract time stamp of this OLE item (this is in GMT)
#
sub _extract_ole_date {
  my ($self, $pps) = @_;
  unless (defined ($self->{OLEDATE})) {
    # Make Date
    my $datearr;
    $datearr = $pps->{Time2nd};
    $datearr = $pps->{Time1st} unless $datearr and $datearr->[0];
    $self->{OLEDATE} = $self->_format_date($datearr) if $datearr and $datearr->[0];
  }
  return;
}

# If we didn't get the date from the original header data, we may be able
# to get it from the SUBMISSION_ID:
# It seems to have the format of a semicolon-separated list of key=value
# pairs. The key l has a value with the format:
# <SERVER>-<DATETIME>Z-<NUMBER>, where DATETIME is the date and time (gmt)
# in the format YYMMDDHHMMSS.
sub _submission_id_date {
  my $self = shift;

  my $submission_id = $self->{SUBMISSION_ID} or return;
  $submission_id =~ m/ l=.*- (\d\d) (\d\d) (\d\d) (\d\d) (\d\d) (\d\d) Z-.* /x
    or return;
  my $year = $1;
  $year += 100 if $year < 20;
  return $self->_format_date([$6,$5,$4,$3,$2-1,$year]);
}

sub _SaveAttachment {
  my ($self, $mime, $att) = @_;

  my $m = $att->to_email_mime;
  $self->_clean_part_header($m);
  $mime->parts_add([$m]);
  return;
}

# Set header fields
sub _AddHeaderField {
  my ($self, $mime, $fieldname, $value) = @_;

  #my $oldvalue = $mime->header($fieldname);
  #return if $oldvalue;
  $mime->header_set($fieldname, $value) if $value;
  return;
}

sub _Address {
  my ($self, $tag) = @_;

  my $result = $self->{$tag} || "";

  my $address = $self->{$tag . "_ADDR"} || "";
  if ($address) {
    $result .= " " if $result;
    $result .= "<$address>";
  }

  return $result;
}

# Find SMTP addresses for the given list of names
sub _expand_address_list {
  my ($self, $names) = @_;

  return "" unless defined $names;

  my @namelist = split / ; [ ]* /x, $names;
  my @result;
  name: foreach my $name (@namelist) {
    my $addresstext = $self->_find_name_in_addresspool($name);
    if ($addresstext) {
      push @result, $addresstext;
    } else {
      push @result, $name;
    }
  }
  return join ", ", @result;
}

sub _find_name_in_addresspool {
  my ($self, $name) = @_;

  my $addresspool = $self->{ADDRESSES};

  foreach my $address (@{$addresspool}) {
    if ($name eq $address->name) {
      return $address->display_address;
    }
  }
  return;
}

# TODO: Don't really want to need this!
sub _clean_part_header {
  my ($self, $part) = @_;
  $part->header_set('Date');
  unless ($part->content_type =~ m{ ^ multipart / }x) {
    $part->header_set('MIME-Version')
  };
  return;
}

sub _create_mime_plain_body {
  my $self = shift;
  return Email::MIME->create(
    attributes => {
      content_type => "text/plain",
      charset => "ISO-8859-1",
      disposition => "inline",
      encoding => "8bit",
    },
    body => $self->{BODY_PLAIN}
  );
}

sub _create_mime_html_body {
  my $self = shift;
  return Email::MIME->create(
    attributes => {
      content_type => "text/html",
      disposition => "inline",
      encoding => "8bit",
    },
    body => $self->{BODY_HTML}
  );
}

# Implementation based on the information in
# http://www.freeutils.net/source/jtnef/rtfcompressed.jsp,
# and the implementation in tnef version 1.4.5.
my $MAGIC_COMPRESSED_RTF = 0x75465a4c;
my $MAGIC_UNCOMPRESSED_RTF = 0x414c454d;
my $BASE_BUFFER =
  "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman "
  . "\\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial"
  . "Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\n\r\\par "
  . "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";


sub _create_mime_rtf_body {
  my $self = shift;
  my $data = $self->{BODY_RTF};

  my ($size, $rawsize, $magic, $crc) = unpack "V4", substr $data, 0, 16;

  my $buffer;

  if ($magic == $MAGIC_COMPRESSED_RTF) {
    $buffer = $BASE_BUFFER;
    my $output_length = length($buffer) + $rawsize;
    my @flags;
    my $in = 16;
    while (length($buffer) < $output_length) {
      if (@flags == 0) {
        @flags = split "", unpack "b8", substr $data, $in++, 1;
      }
      my $flag = shift @flags;
      if ($flag eq "0") {
        $buffer .= substr $data, $in++, 1;
      } else {
        my ($a, $b) = unpack "C2", substr $data, $in, 2;
        my $offset = ($a << 4) | ($b >> 4);
        my $length = ($b & 0xf) + 2;
        my $buflen = length $buffer;
        my $longoffset = $buflen - ($buflen % 4096) + $offset;
        if ($longoffset >= $buflen) { $longoffset -= 4096; }
        while ($length > 0) {
          $buffer .= substr $buffer, $longoffset, 1;
          $length--;
          $longoffset++;
        }
        $in += 2;
      }
    }
    $buffer = substr $buffer, length $BASE_BUFFER;
  } elsif ($magic == $MAGIC_UNCOMPRESSED_RTF) {
    $buffer = substr $data, length $BASE_BUFFER;
  } else {
    carp "Incorrect magic number in RTF body.\n";
  }
  return Email::MIME->create(
    attributes => {
      content_type => "application/rtf",
      disposition => "inline",
      encoding => "base64",
    },
    body => $buffer
  );
}
# Copy original header data.
# Note: This should contain the Date: header.
sub _copy_header_data {
  my ($self, $mime) = @_;

  defined $self->{HEAD} or return;

  # The extra \n is neede for Email::Simple to pick up all headers.
  # This is a change in Email::Simple.
  my $parsed = Email::Simple->new($self->{HEAD} . "\n");

  foreach my $tag (grep { !$skipheaders->{uc $_}} $parsed->header_names) {
    $mime->header_set($tag, $parsed->header($tag));
  }
  return;
}

# Set header fields
sub _SetHeaderFields {
  my ($self, $mime) = @_;

  $self->_AddHeaderField($mime, 'Subject', $self->{SUBJECT});
  $self->_AddHeaderField($mime, 'From', $self->_Address("FROM"));
  #$self->_AddHeaderField($mime, 'Reply-To', $self->_Address("REPLYTO"));
  $self->_AddHeaderField($mime, 'To', $self->_expand_address_list($self->{TO}));
  $self->_AddHeaderField($mime, 'Cc', $self->_expand_address_list($self->{CC}));
  $self->_AddHeaderField($mime, 'Message-Id', $self->{MESSAGEID});
  $self->_AddHeaderField($mime, 'In-Reply-To', $self->{INREPLYTO});
  $self->_AddHeaderField($mime, 'References', $self->{REFERENCES});

  # Least preferred option to set the Date: header; this uses the date the
  # msg file was saved.
  $self->_AddHeaderField($mime, 'Date', $self->{OLEDATE});

  # Second preferred option: get it from the SUBMISSION_ID:
  $self->_AddHeaderField($mime, 'Date', $self->_submission_id_date());

  # Most prefered option from the property list
  $self->_AddHeaderField($mime, 'Date', $self->{DATE2ND});
  $self->_AddHeaderField($mime, 'Date', $self->{DATE1ST});

  # After this, we'll try getting the date from the original headers.
  return;
}

1;

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

Not all data that's in the .msg file is converted. There simply are some
parts whose meaning escapes me. Formatting of text messages will also be
lost. GPG signed mail is not processed correctly.

=head1 AUTHOR

Matijs van Zuijlen, C<matijs@matijs.net>

=head1 COPYRIGHT AND LICENSE

Copyright 2002, 2004, 2006--2009 by Matijs van Zuijlen

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut
use strict;
use warnings;
use Email::Simple;
use Email::MIME::Creator;
use Email::MIME::ContentType;
use OLE::Storage_Lite;
use POSIX;
use Encode;
use Carp;

my $DIR_TYPE = 1;
my $FILE_TYPE = 2;

use vars qw($VERSION);
$VERSION = "0.906";
#
# Descriptions partially based on mapitags.h
#
my $skipproperties = {
  # Envelope properties
  '000B' => "Conversation key?",
  '001A' => "Type of message",
  '003B' => "Sender address variant",
  '003D' => "Contains 'Re: '",
  '003F' => "'recieved by' id",
  '0040' => "'recieved by' name",
  # TODO: These two fields are part of the Sender field.
  '0041' => "Sender variant address id",
  '0042' => "Sender variant name",
  '0043' => "'recieved representing' id",
  '0044' => "'recieved representing' name",
  '0046' => "Read receipt address id",
  '0051' => "'recieved by' search key",
  '0052' => "'recieved representing' search key",
  '0053' => "Read receipt search key",
  # TODO: These two fields are part of the Sender field.
  '0064' => "Sender variant address type",
  '0065' => "Sender variant address",
  '0070' => "Conversation topic",
  '0071' => "Conversation index",
  '0075' => "'recieved by' address type",
  '0076' => "'recieved by' email address",
  '0077' => "'recieved representing' address type",
  '0078' => "'recieved representing' email address",
  '007F' => "something like a message id",
  # Recipient properties
  '0C19' => "Reply address variant",
  '0C1D' => "Reply address variant",
  '0C1E' => "Reply address type",
  # Non-transmittable properties
  '0E02' => "?Should BCC be displayed",
  '0E0A' => "sent mail id",
  '0E1D' => "Subject w/o Re",
  '0E27' => "64 bytes: Unknown",
  '0FF6' => "Index",
  '0FF9' => "Index",
  '0FFF' => "Address variant",
  # Content properties
  '1008' => "Summary or something",
  '1009' => "RTF Compressed",
  # --
  '1046' => "From address variant",
  # 'Common property'
  '3001' => "Display name",
  '3002' => "Address Type",
  '300B' => "'Search key'",
  # Message store info
  '3414' => "Message Store Provider",
  # Attachment properties
  '3702' => "Attachment encoding",
  '3703' => "Attachment extension",
  '3709' => "WMF with attachment rendering info", # Maybe an icon or something?
  '370A' => "Tag identifying application that supplied the attachment",
  '3713' => "Icon URL?",
  # 'Mail user'
  '3A20' => "Address variant",
  # 3900 -- 39FF: 'Address book'
  '39FF' => "7 bit display name",
  # 'Display table properties'
  '3FF8' => "Routing data?",
  '3FF9' => "Routing data?",
  '3FFA' => "Routing data?",
  '3FFB' => "Routing data?",
  # 'Transport-defined envelope property'
  '4029' => "Sender variant address type",
  '402A' => "Sender variant address",
  '402B' => "Sender variant name",
  '5FF6' => "Recipient name",
  '5FF7' => "Recipient address variant",
  # 'Provider-defined internal non-transmittable property'
  '6740' => "Unknown, binary data",
  # User defined id's
  '8000' => "Content Class",
  '8002' => "Unknown, binary data",
};

my $skipheaders = {
  map { uc($_) => 1 }
  "MIME-Version",
  "Content-Type",
  "Content-Transfer-Encoding",
  "X-Mailer",
  "X-Msgconvert",
  "X-MS-Tnef-Correlator",
  "X-MS-Has-Attach"
};

my $ENCODING_UNICODE = '001F';
my $ENCODING_ASCII = '001E';
my $ENCODING_BINARY = '0102';
my $ENCODING_DIRECORY = '000D';

my $KNOWN_ENCODINGS = {
  '000D' => 'Directory',
  '001F' => 'Unicode',
  '001E' => 'Ascii?',
  '0102' => 'Binary',
};

my $MAP_ATTACHMENT_FILE = {
  '3701' => "DATA",        # Data
  '3704' => "SHORTNAME",   # Short file name
  '3707' => "LONGNAME",    # Long file name
  '370E' => "MIMETYPE",    # mime type
  '3712' => "CONTENTID",   # content-id
  '3716' => "DISPOSITION", # disposition
};

my $MAP_SUBITEM_FILE = {
  '1000' => "BODY_PLAIN",      # Body
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
  '1042' => "INREPLYTO",       # In reply to Message-Id
};

my $MAP_ADDRESSITEM_FILE = {
  '3001' => "NAME",	    # Real name
  '3002' => "TYPE",         # Address type
  '403D' => "TYPE",         # Address type
  '3003' => "ADDRESS",      # Address
  '403E' => "ADDRESS",      # Address
  '39FE' => "SMTPADDRESS",  # SMTP Address variant
};

#
# Main body of module
#

sub new {
  my $class = shift;
  my $file = shift or croak "File name is required parameter";
  my $verbose = shift;

  my $self = $class->_empty_new;

  my $msg = OLE::Storage_Lite->new($file);
  my $pps = $msg->getPpsTree(1);
  $pps or croak "Parsing $file as OLE file failed";
  $self->_set_verbosity($verbose);
  $self->_process_root_dir($pps);

  return $self;
}

sub _empty_new {
  my $class = shift;

  return bless {
    ADDRESSES => [], ATTACHMENTS => [], FROM_ADDR_TYPE => "",
    HAS_UNICODE => 0, VERBOSE => 0,
  }, $class;
}

sub to_email_mime {
  my $self = shift;

  my ($plain, $html);
  my $bodymime;
  my $mime;

  if ($self->{BODY_PLAIN} or not $self->{BODY_HTML}) {
    $plain = $self->_create_mime_plain_body();
  }
  if ($self->{BODY_HTML}) {
    $html = $self->_create_mime_html_body();
  }

  if ($html and $plain) {
    $self->_clean_part_header($plain);
    $self->_clean_part_header($html);
    $bodymime = Email::MIME->create(
      attributes => {
	content_type => "multipart/alternative",
	encoding => "8bit",
      },
      parts => [$plain, $html]
    );
  } elsif ($html) {
    $bodymime = $html;
  } else {
    $bodymime = $plain;
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

sub _set_verbosity {
  my ($self, $verbosity) = @_;
  $self->{VERBOSE} = $verbosity ? 1 : 0;
  return;
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

#
# _process_root_dir: Check Root Entry, parse sub-entries.
# The OLE file consists of a single entry called Root Entry, which has
# several children. These children are parsed in the sub SubItem.
# 
sub _process_root_dir {
  my ($self, $pps) = @_;

  foreach my $child (@{$pps->{Child}}) {
    if ($child->{Type} == $DIR_TYPE) {
      $self->_process_subdirectory($child);
    } elsif ($child->{Type} == $FILE_TYPE) {
      $self->_process_pps_file_entry($child, $self, $MAP_SUBITEM_FILE);
    } else {
      carp "Unknown entry type: $child->{Type}";
    }
  }
  return;
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

  my $addr_info = { NAME => undef, ADDRESS => undef, TYPE => "" };

  foreach my $child (@{$pps->{Child}}) {
    if ($child->{Type} == $DIR_TYPE) {
      $self->_warn_about_unknown_directory($child); # DIR Entries: There should be none.
    } elsif ($child->{Type} == $FILE_TYPE) {
      $self->_process_pps_file_entry($child, $addr_info, $MAP_ADDRESSITEM_FILE);
    } else {
      carp "Unknown entry type: $child->{Type}";
    }
  }
  push @{$self->{ADDRESSES}}, $addr_info;
  return;
}

#
# Process a subdirectory that contains an attachment.
#
sub _process_attachment {
  my ($self, $pps) = @_;

  my $attachment = {
    SHORTNAME   => undef,
    LONGNAME    => undef,
    MIMETYPE    => 'application/octet-stream',
    ENCODING    => 'base64',
    DISPOSITION => 'attachment',
    CONTENTID   => undef,
    DATA        => undef,
  };
  foreach my $child (@{$pps->{Child}}) {
    if ($child->{Type} == $DIR_TYPE) {
      $self->_process_attachment_subdirectory($child, $attachment);
    } elsif ($child->{Type} == $FILE_TYPE) {
      $self->_process_pps_file_entry($child, $attachment, $MAP_ATTACHMENT_FILE);
    } else {
      carp "Unknown entry type: $child->{Type}";
    }
  }
  if ($attachment->{MIMETYPE} eq 'multipart/signed') {
    $attachment->{ENCODING} = '8bit';
  }
  push @{$self->{ATTACHMENTS}}, $attachment;
  return;
}

#
# Process a subdirectory that is part of an attachment
#
sub _process_attachment_subdirectory {
  my ($self, $pps, $att) = @_;
  my $name = $self->_get_pps_name($pps);
  my ($property, $encoding) = $self->_parse_item_name($name);

  if ($property eq '3701') { # Nested msg file
    my $msgp = ref($self)->_empty_new();
    $msgp->_set_verbosity($self->{VERBOSE});
    $msgp->_process_root_dir($pps);

    $att->{DATA} = $msgp->to_email_mime->as_string;
    $att->{MIMETYPE} = 'message/rfc822';
    $att->{ENCODING} = '8bit';
  } else {
    $self->_warn_about_unknown_directory($pps);
  }
  return;
}

#
# Generic processor for a file entry: Inserts the entry's data into the
# hash $target, using the $map to find the proper key.
# TODO: Mapping should probably be applied at a later time instead.
#
sub _process_pps_file_entry {
  my ($self, $pps, $target, $map) = @_;

  my $name = $self->_get_pps_name($pps);
  my ($property, $encoding) = $self->_parse_item_name($name);

  if (defined $property and my $key = $map->{$property}) {
    my $data = $pps->{Data};
    if ($encoding eq $ENCODING_DIRECORY) {
      die "Unexpected directory encoding for property $name";
    }
    if ($encoding ne $ENCODING_BINARY) {
      if ($encoding eq $ENCODING_UNICODE) {
	$data = decode("UTF-16LE", $data);
      }
      $data =~ s/\000$//sg;
      $data =~ s/\r\n/\n/sg;
    }
    $target->{$key} = $data;
  } else {
    $self->_warn_about_unknown_file($pps);
  }
  return;
}

sub _warn_about_unknown_directory {
  my ($self, $pps) = @_;

  my $name = $self->_get_pps_name($pps);
  if ($name eq '__nameid_version1 0') {
    # TODO: Use this data to access so-called named properties.
    $self->{VERBOSE}
      and warn "Skipping DIR entry $name (Introductory stuff)\n";
  } else {
    warn "Unknown DIR entry $name\n";
  }
  return;
}

sub _warn_about_unknown_file {
  my ($self, $pps) = @_;

  my $name = $self->_get_pps_name($pps);

  if ($name eq '__properties_version1 0'
      or $name eq 'Olk10SideProps_0001') {
    $self->{VERBOSE}
      and warn "Skipping FILE entry $name (Properties)\n";
    return;
  }

  # FIXME: encoding not used.
  my ($property, $encoding) = $self->_parse_item_name($name);
  unless (defined $property) {
    warn "Unknown FILE entry $name\n";
    return;
  }
  if ($skipproperties->{$property}) {
    $self->{VERBOSE}
      and warn "Skipping property $property ($skipproperties->{$property})\n";
  } elsif (not $self->_is_transmittable_property($property)) {
    $self->{VERBOSE}
      and warn "Skipping property $property (non-transmittable property)\n";
  } elsif ($property =~ /^80/) {
    $self->{VERBOSE}
      and warn "Skipping property $property (user-defined property)\n";
  } elsif ($pps->{Data} eq "") {
    $self->{VERBOSE}
      and warn "Unknown property $property (no data)\n";
  } else {
    warn "Unknown property $property\n";
  }
  return;
}

#
# Helper functions
#

sub _is_transmittable_property {
  my ($self, $prop) = @_;
  return 1 if $prop lt '0E00';
  return 1 if $prop ge '1000' and $prop lt '6000';
  return 1 if $prop ge '6800' and $prop lt '7C00';
  return 1 if $prop ge '8000';
  return 0;
}

sub _get_pps_name {
  my ($self, $pps) = @_;
  my $name = OLE::Storage_Lite::Ucs2Asc($pps->{Name});
  $name =~ s/\W/ /g;
  return $name;
}

#
# Extract time stamp of this OLE item (this is in GMT)
#
sub _extract_ole_date {
  my ($self, $pps) = @_;
  unless (defined ($self->{OLEDATE})) {
    # Make Date
    my $datearr;
    $datearr = $pps->{Time2nd};
    $datearr = $pps->{Time1st} unless($datearr);
    $self->{OLEDATE} = $self->_format_date($datearr) if $datearr;
  }
  return;
}

#
# Format a gmt date according to RFC822
#
sub _format_date {
  my ($self, $datearr) = @_;
  my $day = qw(Sun Mon Tue Wed Thu Fri Sat)[strftime("%w", @$datearr)];
  my $month = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec)[strftime("%m", @$datearr) - 1];
  return strftime("$day, %d $month %Y %H:%M:%S +0000", @$datearr);
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
  $submission_id =~ m/l=.*-(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)Z-.*/
    or return;
  my $year = $1;
  $year += 100 if $year < 20;
  return $self->_format_date([$6,$5,$4,$3,$2-1,$year]);
}

sub _parse_item_name {
  my ($self, $name) = @_;

  if ($name =~ /^__substg1 0_(....)(....)$/) {
    my ($property, $encoding) = ($1, $2);
    if ($encoding eq $ENCODING_UNICODE and not ($self->{HAS_UNICODE})) {
      $self->{HAS_UNICODE} = 1;
    } elsif (not $KNOWN_ENCODINGS->{$encoding}) {
      warn "Unknown encoding $encoding. Results may be strange or wrong.\n";
    }
    return ($property, $encoding);
  } else {
    return (undef, undef);
  }
}

sub _SaveAttachment {
  my ($self, $mime, $att) = @_;

  my $mt = parse_content_type($att->{MIMETYPE});
  my $m = Email::MIME->create(
    attributes => {
      content_type => "$mt->{discrete}/$mt->{composite}",
      %{$mt->{attributes}},
      encoding => $att->{ENCODING},
      filename => ($att->{LONGNAME} ? $att->{LONGNAME} : $att->{SHORTNAME}),
      name => ($att->{LONGNAME} ? $att->{LONGNAME} : $att->{SHORTNAME}),
      disposition => $att->{DISPOSITION},
    },
    header => [ 'Content-ID' => $att->{CONTENTID} ],
    body => $att->{DATA});
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
sub _ExpandAddressList {
  my ($self, $names) = @_;

  return "" unless defined $names;

  my $addresspool = $self->{ADDRESSES};
  my @namelist = split /; */, $names;
  my @result;
  name: foreach my $name (@namelist) {
    foreach my $address (@$addresspool) {
      if ($name eq $address->{NAME}) {
	my $addresstext = $address->{NAME} . " <";
	if (defined ($address->{SMTPADDRESS})) {
	  $addresstext .= $address->{SMTPADDRESS};
	} elsif ($address->{TYPE} eq "SMTP") {
	  $addresstext .= $address->{ADDRESS};
	}
	$addresstext .= ">";
	push @result, $addresstext;
	next name;
      }
    }
    push @result, $name;
  }
  return join ", ", @result;
}

# TODO: Don't really want to need this!
sub _clean_part_header {
  my ($self, $part) = @_;
  $part->header_set('Date');
  unless ($part->content_type =~ /^multipart\//) {
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
  my $body = $self->{BODY_HTML};
  # FIXME: This makes sure tests succeed for now, but is not really
  # necessary for correct display in the mail reader.
  $body =~ s/\r\n/\n/sg;
  return Email::MIME->create(
    attributes => {
      content_type => "text/html",
      disposition => "inline",
      encoding => "8bit",
    },
    body => $body
  );
}
# Copy original header data.
# Note: This should contain the Date: header.
sub _copy_header_data {
  my ($self, $mime) = @_;

  defined $self->{HEAD} or return;

  # The extra \n is neede for Email::Simple to pick up all headers.
  # This is a change in Email::Simple.
  my $parsed = new Email::Simple($self->{HEAD} . "\n");

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
  $self->_AddHeaderField($mime, 'To', $self->_ExpandAddressList($self->{TO}));
  $self->_AddHeaderField($mime, 'Cc', $self->_ExpandAddressList($self->{CC}));
  $self->_AddHeaderField($mime, 'Message-Id', $self->{MESSAGEID});
  $self->_AddHeaderField($mime, 'In-Reply-To', $self->{INREPLYTO});

  # Least preferred option to set the Date: header; this uses the date the
  # msg file was saved.
  $self->_AddHeaderField($mime, 'Date', $self->{OLEDATE});

  # Second preferred option: get it from the SUBMISSION_ID:
  $self->_AddHeaderField($mime, 'Date', $self->_submission_id_date());

  # After this, we'll try getting the date from the original headers.
  return;
}

1;

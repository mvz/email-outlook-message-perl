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

Copyright 2002, 2004, 2006--2009 by Matijs van Zuijlen

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
use strict;
use warnings;
use 5.006;
use vars qw($VERSION);
$VERSION = "0.910";

package Email::Outlook::Message::Base;
use strict;
use warnings;
use Encode;
use IO::String;
use POSIX;
use Carp;
use OLE::Storage_Lite;

my $DIR_TYPE = 1;
my $FILE_TYPE = 2;

# Variable encodings
my $ENCODING_UNICODE = '001F';
my $ENCODING_ASCII = '001E';
my $ENCODING_BINARY = '0102';
my $ENCODING_DIRECTORY = '000D';

our $VARIABLE_ENCODINGS = {
  '000D' => 'Directory',
  '001F' => 'Unicode',
  '001E' => 'Ascii?',
  '0102' => 'Binary',
};

# Fixed encodings
my $ENCODING_INTEGER16 = '0002';
my $ENCODING_INTEGER32 = '0003';
my $ENCODING_BOOLEAN = '000B';
my $ENCODING_DATE = '0040';

#
# Descriptions partially based on mapitags.h
#
our $skipproperties = {
  # Envelope properties
  '0002' => "Alternate Recipient Allowed",
  '000B' => "Conversation Key",
  '0017' => "Importance", #TODO: Use this.
  '001A' => "Message Class",
  '0023' => "Originator Delivery Report Requested",
  '0026' => "Priority", #TODO: Use this.
  '0029' => "Read Receipt Requested", #TODO: Use this.
  '0036' => "Sensitivity", # As assessed by the Sender
  '003B' => "Sent Representing Search Key",
  '003D' => "Subject Prefix",
  '003F' => "Received By EntryId",
  '0040' => "Received By Name",
  # TODO: These two fields are part of the Sender field.
  '0041' => "Sent Representing EntryId",
  '0042' => "Sent Representing Name",
  '0043' => "Received Representing EntryId",
  '0044' => "Received Representing Name",
  '0046' => "Read Receipt EntryId",
  '0051' => "Received By Search Key",
  '0052' => "Received Representing Search Key",
  '0053' => "Read Receipt Search Key",
  # TODO: These two fields are part of the Sender field.
  '0064' => "Sent Representing Address Type",
  '0065' => "Sent Representing Email Address",
  '0070' => "Conversation Topic",
  '0071' => "Conversation Index",
  '0075' => "Received By Address Type",
  '0076' => "Received By Email Address",
  '0077' => "Received Representing Address Type",
  '0078' => "Received Representing Email Address",
  '007F' => "TNEF Correlation Key",
  # Recipient properties
  '0C15' => "Recipient Type",
  # Sender properties
  '0C19' => "Sender Entry Id",
  '0C1D' => "Sender Search Key",
  '0C1E' => "Sender Address Type",
  # Non-transmittable properties
  '0E02' => "Display Bcc",
  '0E06' => "Message Delivery Time",
  '0E07' => "Message Flags",
  '0E0A' => "Sent Mail EntryId",
  '0E0F' => "Responsibility",
  '0E1B' => "Has Attachments",
  '0E1D' => "Normalized Subject",
  '0E1F' => "RTF In Sync",
  '0E20' => "Attachment Size",
  '0E21' => "Attachment Number",
  '0E23' => "Internet Article Number",
  '0E27' => "Security Descriptor",
  '0E79' => "Trust Sender",
  '0FF4' => "Access",
  '0FF6' => "Instance Key",
  '0FF7' => "Access Level",
  '0FF9' => "Record Key",
  '0FFE' => "Object Type",
  '0FFF' => "EntryId",
  # Content properties
  '1006' => "RTF Sync Body CRC",
  '1007' => "RTF Sync Body Count",
  '1008' => "RTF Sync Body Tag",
  '1010' => "RTF Sync Prefix Count",
  '1011' => "RTF Sync Trailing Count",
  '1046' => "Original Message ID",
  '1080' => "Icon Index",
  '1081' => "Last Verb Executed",
  '1082' => "Last Verb Execution Time",
  '10F3' => "URL Component Name",
  '10F4' => "Attribute Hidden",
  '10F5' => "Attribute System",
  '10F6' => "Attribute Read Only",
  # 'Common property'
  '3000' => "Row Id",
  '3001' => "Display Name",
  '3002' => "Address Type",
  '3007' => "Creation Time",
  '3008' => "Last Modification Time",
  '300B' => "Search Key",
  # Message store info
  '340D' => "Store Support Mask",
  '3414' => "Message Store Provider",
  # Attachment properties
  '3702' => "Attachment Encoding",
  '3703' => "Attachment Extension",
  # TODO: Use the following to distinguish between nested msg and other OLE
  # stores.
  '3705' => "Attachment Method",
  '3709' => "Attachment Rendering", # Icon as WMF
  '370A' => "Tag identifying application that supplied the attachment",
  '370B' => "Attachment Rendering Position",
  '3713' => "Attachment Content Location", #TODO: Use this?
  # 3900 -- 39FF: 'Address book'
  '3900' => "Address Book Display Type",
  '39FF' => "Address Book 7 Bit Display Name",
  # Mail User Object
  '3A00' => "Account",
  '3A20' => "Transmittable Display Name",
  '3A40' => "Send Rich Info",
  '3FDE' => "Internet Code Page", # TODO: Perhaps use this.
  # 'Display table properties'
  '3FF8' => "Creator Name",
  '3FF9' => "Creator EntryId",
  '3FFA' => "Last Modifier Name",
  '3FFB' => "Last Modifier EntryId",
  '3FFD' => "Message Code Page",
  # 'Transport-defined envelope property'
  '4019' => "Sender Flags",
  '401A' => "Sent Representing Flags",
  '401B' => "Received By Flags",
  '401C' => "Received Representing Flags",
  '4029' => "Read Receipt Address Type",
  '402A' => "Read Receipt Email Address",
  '402B' => "Read Receipt Name",
  '5FF6' => "Recipient Display Name",
  '5FF7' => "Recipient EntryId",
  '5FFD' => "Recipient Flags",
  '5FFF' => "Recipient Track Status",
  # 'Provider-defined internal non-transmittable property'
  '664A' => "Has Named Properties",
  '6740' => "Sent Mail Server EntryId",
};

sub new {
  my ($class, $pps, $verbose) = @_;
  my $self = bless {
    _pps_file_entries => {},
    _pps => $pps
  }, $class;
  $self->_set_verbosity($verbose);
  $self->_process_pps($pps);
  return $self;
}

sub mapi_property_names {
  my $self = shift;
  return keys %{$self->{_pps_file_entries}};
}

sub get_mapi_property {
  my ($self, $code) = @_;
  return $self->{_pps_file_entries}->{$code};
}

sub set_mapi_property {
  my ($self, $code, $data) = @_;
  $self->{_pps_file_entries}->{$code} = $data;
  return;
}

sub property {
  my ($self, $name) = @_;
  my $map = $self->_property_map;
  # TODO: Prepare reverse map instead of doing dumb lookup.
  foreach my $code (keys %$map) {
    my $key = $map->{$code};
    next unless $key eq $name;
    my $prop = $self->get_mapi_property($code);
    if ($prop) {
      my ($encoding, $data) = @{$prop};
      return $self->_decode_mapi_property($encoding, $data);
    } else {
      return;
    }
  }
  return;
}

sub _decode_mapi_property {
  my ($self, $encoding, $data) = @_;
  if ($encoding eq $ENCODING_ASCII or $encoding eq $ENCODING_UNICODE) {
    if ($encoding eq $ENCODING_UNICODE) {
      $data = decode("UTF-16LE", $data);
    }
    $data =~ s/\000$//sg;
    $data =~ s/\r\n/\n/sg;
    return $data
  } elsif ($encoding eq $ENCODING_BINARY) {
    return $data
  } elsif ($encoding eq $ENCODING_DATE) {
    my @a = OLE::Storage_Lite::OLEDate2Local $data;
    return $self->_format_date(\@a);
  } elsif ($encoding eq $ENCODING_INTEGER16) {
    return unpack("v", substr($data, 0, 2));
  } elsif ($encoding eq $ENCODING_INTEGER32) {
    return unpack("V", substr($data, 0, 4));
  } elsif ($encoding eq $ENCODING_BOOLEAN) {
    return unpack("C", substr($data, 0, 1));
  }

  die "Unexpected encoding $encoding";
}

sub _process_pps {
  my ($self, $pps) = @_;
  foreach my $child (@{$pps->{Child}}) {
    if ($child->{Type} == $DIR_TYPE) {
      $self->_process_subdirectory($child);
    } elsif ($child->{Type} == $FILE_TYPE) {
      $self->_process_pps_file_entry($child);
    } else {
      carp "Unknown entry type: $child->{Type}";
    }
  }
  $self->_check_pps_file_entries($self->_property_map);
  return;
}

sub _get_pps_name {
  my ($self, $pps) = @_;
  my $name = OLE::Storage_Lite::Ucs2Asc($pps->{Name});
  $name =~ s/\W/ /g;
  return $name;
}

sub _parse_item_name {
  my ($self, $name) = @_;

  if ($name =~ /^__substg1 0_(....)(....)$/) {
    my ($property, $encoding) = ($1, $2);
    if ($encoding eq $ENCODING_UNICODE and not ($self->{HAS_UNICODE})) {
      $self->{HAS_UNICODE} = 1;
    } elsif (not $VARIABLE_ENCODINGS->{$encoding}) {
      warn "Unknown encoding $encoding. Results may be strange or wrong.\n";
    }
    return ($property, $encoding);
  } else {
    return (undef, undef);
  }
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
  } else {
    warn "Unknown FILE entry $name\n";
  }
  return;
}

#
# Generic processor for a file entry: Inserts the entry's data into the
# $self's mapi property list.
#
sub _process_pps_file_entry {
  my ($self, $pps) = @_;
  my $name = $self->_get_pps_name($pps);
  my ($property, $encoding) = $self->_parse_item_name($name);

  if (defined $property) {
    $self->set_mapi_property($property, [$encoding, $pps->{Data}]);
  } elsif ($name eq '__properties_version1 0') {
    $self->_process_property_stream ($pps->{Data});
  } else {
    $self->_warn_about_unknown_file($pps);
  }
  return;
}

sub _process_property_stream {
  my ($self, $data) = @_;
  my ($n, $len) = ($self->_property_stream_header_length, length $data) ;

  while ($n + 16 <= $len) {
    my @f = unpack "v4", substr $data, $n, 8;

    my $encoding = sprintf("%04X", $f[0]);

    unless ($VARIABLE_ENCODINGS->{$encoding}) {
      my $property = sprintf("%04X", $f[1]);
      my $propdata = substr $data, $n+8, 8;
      $self->set_mapi_property($property, [$encoding, $propdata]);
    }
  } continue {
    $n += 16 ;
  }
  return;
}

sub _check_pps_file_entries {
  my ($self, $map) = @_;

  foreach my $property ($self->mapi_property_names) {
    if (my $key = $map->{$property}) {
      $self->_use_property($key, $property);
    } else {
      $self->_warn_about_skipped_property($property);
    }
  }
}

sub _use_property {
  my ($self, $key, $property) = @_;
  my ($encoding, $data) = @{$self->get_mapi_property($property)};
  $self->{$key} = $self->_decode_mapi_property($encoding, $data);

  $self->{VERBOSE}
    and $self->_log_property("Using   ", $property, $encoding, $key, $self->{$key});
}

sub _warn_about_skipped_property {
  my ($self, $property) = @_;

  return unless $self->{VERBOSE};

  my ($encoding, $data) = @{$self->get_mapi_property($property)};
  my $value = $self->_decode_mapi_property($encoding, $data);
  my $meaning = $skipproperties->{$property} || "UNKNOWN";

  $self->_log_property("Skipping", $property, $encoding, $meaning, $value);
  return;
}

sub _log_property {
  my ($self, $message, $property, $encoding, $meaning, $value) = @_;

  $value = substr($value, 0, 50);

  if ($encoding eq $ENCODING_BINARY) {
    if ($value =~ /[[:print:]]/) {
      $value =~ s/[^[:print:]]/./g;
    } else {
      $value =~ s/./sprintf("%02x ", ord($&))/sge;
    }
  }

  if (length($value) > 45) {
    $value = substr($value, 0, 41) . " ...";
  }

  warn "$message property $encoding:$property ($meaning): $value\n";
}

sub _set_verbosity {
  my ($self, $verbosity) = @_;
  $self->{VERBOSE} = $verbosity ? 1 : 0;
  return;
}

sub _is_transmittable_property {
  my ($self, $prop) = @_;
  return 1 if $prop lt '0E00';
  return 1 if $prop ge '1000' and $prop lt '6000';
  return 1 if $prop ge '6800' and $prop lt '7C00';
  return 1 if $prop ge '8000';
  return 0;
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

package Email::Outlook::Message::AddressInfo;
use strict;
use warnings;
use Carp;
use base 'Email::Outlook::Message::Base';

our $MAP_ADDRESSITEM_FILE = {
  '3001' => "NAME",          # Real name
  '3002' => "TYPE",          # Address type
  '403D' => "TYPE2",         # Address type TODO: Not used
  '3003' => "ADDRESS",       # Address
  '403E' => "ADDRESS2",      # Address TODO: Not used
  '39FE' => "SMTPADDRESS",   # SMTP Address variant
};

sub _property_map {
  return $MAP_ADDRESSITEM_FILE;
}

# DIR Entries: There should be none.
sub _process_subdirectory {
  my ($self, $pps) = @_;
  $self->_warn_about_unknown_directory($pps);
}

sub name { return $_[0]->property('NAME') }
sub address_type { return $_[0]->property('TYPE') }
sub address { return $_[0]->property('ADDRESS') }
sub smtp_address { return $_[0]->property('SMTPADDRESS') }

sub display_address {
  my $self = shift;
  my $addresstext = $self->name . " <";
  if (defined ($self->smtp_address)) {
    $addresstext .= $self->smtp_address;
  } elsif ($self->address_type eq "SMTP") {
    $addresstext .= $self->address;
  }
  $addresstext .= ">";
  return $addresstext;
}

sub _property_stream_header_length { 8 }

package Email::Outlook::Message::Attachment;
use strict;
use warnings;
use Carp;
use Email::MIME::ContentType;
use base 'Email::Outlook::Message::Base';

our $MAP_ATTACHMENT_FILE = {
  '3701' => "DATA",        # Data
  '3704' => "SHORTNAME",   # Short file name
  '3707' => "LONGNAME",    # Long file name
  '370E' => "MIMETYPE",    # mime type
  '3712' => "CONTENTID",   # content-id
  '3716' => "DISPOSITION", # disposition
};

sub new {
  my ($class, $pps, $verbosity) = @_;
  my $self = $class->SUPER::new($pps, $verbosity);
  bless $self, $class;
  $self->{MIMETYPE} ||= 'application/octet-stream';
  $self->{ENCODING} ||= 'base64';
  $self->{DISPOSITION} ||= 'attachment';
  if ($self->{MIMETYPE} eq 'multipart/signed') {
    $self->{ENCODING} = '8bit';
  }
  return $self;
}

sub to_email_mime {
  my $self = shift;

  my $mt = parse_content_type($self->{MIMETYPE});
  my $m = Email::MIME->create(
    attributes => {
      content_type => "$mt->{discrete}/$mt->{composite}",
      %{$mt->{attributes}},
      encoding => $self->{ENCODING},
      filename => $self->{LONGNAME} || $self->{SHORTNAME},
      name => $self->{LONGNAME} || $self->{LONGNAME},
      disposition => $self->{DISPOSITION},
    },
    header => [ 'Content-ID' => $self->{CONTENTID} ],
    body => $self->{DATA});
  return $m
}

sub _property_map {
  return $MAP_ATTACHMENT_FILE;
}

sub _process_subdirectory {
  my ($self, $pps) = @_;
  my $name = $self->_get_pps_name($pps);
  my ($property, $encoding) = $self->_parse_item_name($name);

  if ($property eq '3701') { # Nested msg file
    my $is_msg = 1;
    foreach my $child (@{$pps->{Child}}) {
      my $name = $self->_get_pps_name($child);
      unless (
	$name =~ /^__recip/ or $name =~ /^__attach/
	  or $name =~ /^__substg1/ or $name =~ /^__nameid/
	  or $name =~ /^__properties/
      ) {
	$is_msg = 0;
	last;
      }
    }
    if ($is_msg) {
      my $msgp = Email::Outlook::Message->_empty_new();
      $msgp->_set_verbosity($self->{VERBOSE});
      $msgp->_process_pps($pps);

      $self->{DATA} = $msgp->to_email_mime->as_string;
      $self->{MIMETYPE} = 'message/rfc822';
      $self->{ENCODING} = '8bit';
    } else {
      foreach my $child (@{$pps->{Child}}) {
	if ($child->{Type} == $FILE_TYPE) {
	  foreach my $prop ("Time1st", "Time2nd") {
	    $child->{$prop} = undef;
	  }
	}
      }
      my $nPps = OLE::Storage_Lite::PPS::Root->new(
	$pps->{Time1st}, $pps->{Time2nd}, $pps->{Child});
      my $data;
      my $io = IO::String->new($data);
      binmode($io);
      $nPps->save($io, 1);
      $self->{DATA} = $data;
      #      $att->{MIMETYPE} = 'message/rfc822';
      #	    $att->{ENCODING} = '8bit';
    }
  } else {
    $self->_warn_about_unknown_directory($pps);
  }
  return;
}

sub _property_stream_header_length { 8 }

package Email::Outlook::Message;
use strict;
use warnings;
use Email::Simple;
use Email::MIME::Creator;
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

  my ($plain, $html);
  my $bodymime;
  my $mime;

  my @parts;

  if ($self->{BODY_PLAIN}) { push(@parts, $self->_create_mime_plain_body()); }
  if ($self->{BODY_HTML}) { push(@parts, $self->_create_mime_html_body()); }
  if ($self->{BODY_RTF}) { push(@parts, $self->_create_mime_rtf_body()); }

  if ((scalar @parts) > 1) {
    map { $self->_clean_part_header($_) } @parts;

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

  my $addr_info = new Email::Outlook::Message::AddressInfo($pps,
    $self->{VERBOSE});

  push @{$self->{ADDRESSES}}, $addr_info;
  return;
}

#
# Process a subdirectory that contains an attachment.
#
sub _process_attachment {
  my ($self, $pps) = @_;

  my $attachment = new Email::Outlook::Message::Attachment($pps,
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
    $datearr = $pps->{Time1st} unless $datearr and @$datearr[0];
    $self->{OLEDATE} = $self->_format_date($datearr) if $datearr and @$datearr[0];
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
  $submission_id =~ m/l=.*-(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)Z-.*/
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

  my @namelist = split /; */, $names;
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

  foreach my $address (@$addresspool) {
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

# Implementation based on the information in
# http://www.freeutils.net/source/jtnef/rtfcompressed.jsp,
# and the implementation in tnef version 1.4.5.
use constant MAGIC_COMPRESSED_RTF => 0x75465a4c;
use constant MAGIC_UNCOMPRESSED_RTF => 0x414c454d;
use constant BASE_BUFFER =>
  "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman "
  . "\\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial"
  . "Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\n\r\\par "
  . "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";


sub _create_mime_rtf_body {
  my $self = shift;
  my $data = $self->{BODY_RTF};

  my ($size, $rawsize, $magic, $crc) = unpack "V4", substr $data, 0, 16;

  my $buffer;

  if ($magic == MAGIC_COMPRESSED_RTF) {
    $buffer = BASE_BUFFER;
    my $output_length = length($buffer) + $rawsize;
    my @flags;
    my $in = 16;
    while (length($buffer) < $output_length) {
      if (@flags == 0) {
	@flags = split "", unpack "b8", substr $data, $in++, 1;
      }
      my $flag = shift @flags;
      if ($flag == "0") {
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
    $buffer = substr $buffer, length BASE_BUFFER;
  } elsif ($magic == MAGIC_UNCOMPRESSED_RTF) {
    $buffer = substr $data, length BASE_BUFFER;
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

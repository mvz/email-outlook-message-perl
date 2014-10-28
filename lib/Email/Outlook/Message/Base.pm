package Email::Outlook::Message::Base;
=head1 NAME

Email::Outlook::Message::Base - Base parser for .msg files.

=head1 DESCRIPTION

This is an internal module of Email::Outlook::Message.

=head1 METHODS

=over 8

=item B<new($pps, $verbose)>

=item B<get_mapi_property($code)>

=item B<set_mapi_property($code, $data)>

=item B<mapi_property_names()>

=item B<property($name)>

=back

=head1 AUTHOR

Matijs van Zuijlen, C<matijs@matijs.net>

=head1 COPYRIGHT AND LICENSE

Copyright 2002--2014 by Matijs van Zuijlen

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
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
  foreach my $code (keys %{$map}) {
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
    $data =~ s/ \000 $ //sgx;
    return $data;
  }

  if ($encoding eq $ENCODING_BINARY) {
    return $data;
  }

  if ($encoding eq $ENCODING_DATE) {
    my @a = OLE::Storage_Lite::OLEDate2Local $data;
    return $self->_format_date(\@a);
  }

  if ($encoding eq $ENCODING_INTEGER16) {
    return unpack("v", substr($data, 0, 2));
  }

  if ($encoding eq $ENCODING_INTEGER32) {
    return unpack("V", substr($data, 0, 4));
  }

  if ($encoding eq $ENCODING_BOOLEAN) {
    return unpack("C", substr($data, 0, 1));
  }

  warn "Unhandled encoding $encoding\n";
  return $data;
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
  $name =~ s/ \W / /gx;
  return $name;
}

sub _parse_item_name {
  my ($self, $name) = @_;

  if ($name =~ / ^ __substg1 [ ] 0_ (....) (....) $ /x) {
    my ($property, $encoding) = ($1, $2);
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

  if ($name eq 'Olk10SideProps_0001') {
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
  return;
}

sub _use_property {
  my ($self, $key, $property) = @_;
  my ($encoding, $data) = @{$self->get_mapi_property($property)};
  $self->{$key} = $self->_decode_mapi_property($encoding, $data);

  $self->{VERBOSE}
    and $self->_log_property("Using   ", $property, $key);
  return;
}

sub _warn_about_skipped_property {
  my ($self, $property) = @_;

  return unless $self->{VERBOSE};

  my $meaning = $skipproperties->{$property} || "UNKNOWN";

  $self->_log_property("Skipping", $property, $meaning);
  return;
}

sub _log_property {
  my ($self, $message, $property, $meaning) = @_;

  my ($encoding, $data) = @{$self->get_mapi_property($property)};
  my $value = $self->_decode_mapi_property($encoding, $data);
  $value = substr($value, 0, 50);

  if ($encoding eq $ENCODING_BINARY) {
    if ($value =~ / [[:print:]] /x) {
      $value =~ s/ [^[:print:]] /./gx;
    } else {
      $value =~ s/ . / sprintf("%02x ", ord($&)) /sgex;
    }
  }

  if (length($value) > 45) {
    $value = substr($value, 0, 41) . " ...";
  }

  warn "$message property $encoding:$property ($meaning): $value\n";
  return;
}

sub _set_verbosity {
  my ($self, $verbosity) = @_;
  $self->{VERBOSE} = $verbosity ? 1 : 0;
  return;
}

#
# Format a gmt date according to RFC822
#
sub _format_date {
  my ($self, $datearr) = @_;
  my $day = qw(Sun Mon Tue Wed Thu Fri Sat)[strftime("%w", @{$datearr})];
  my $month = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec)[strftime("%m", @{$datearr}) - 1];
  return strftime("$day, %d $month %Y %H:%M:%S +0000", @{$datearr});
}

1;

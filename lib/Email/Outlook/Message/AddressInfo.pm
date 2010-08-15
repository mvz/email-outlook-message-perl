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
  return;
}

sub name { my $self = shift; return $self->property('NAME') }
sub address_type { my $self = shift; return $self->property('TYPE') }
sub address { my $self = shift; return $self->property('ADDRESS') }
sub smtp_address { my $self = shift; return $self->property('SMTPADDRESS') }

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

sub _property_stream_header_length { return 8; }

1;

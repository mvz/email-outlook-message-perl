package Email::Outlook::Message::AddressInfo;
=head1 NAME

Email::Outlook::Message::AddressInfo - Handle addres data in .msg files

=head1 NAME

This is an internal module of Email::Outlook::Message. It is a subclass of
Email::Outlook::Message::Base.

=head1 METHODS

=over 8

=item B<address()>

=item B<address_type()>

=item B<display_address()>

=item B<name()>

=item B<smtp_address()>

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
use vars qw($VERSION);
$VERSION = "0.915";
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

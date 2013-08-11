package Email::Outlook::Message::Attachment;
=head1 NAME

Email::Outlook::Message::Attachment - Handle attachment data in .msg files

=head1 NAME

This is an internal module of Email::Outlook::Message. It is a subclass of
Email::Outlook::Message::Base.

=head1 METHODS

=over 8

=item B<new($pps, $verbosity)>

Create a new attachment object, using $pps as data source. Overrides the base
method by setting some default values.

=item B<to_email_mime()>

Convert the attachment to an Email::MIME object.

=back

=head1 AUTHOR

Matijs van Zuijlen, C<matijs@matijs.net>

=head1 COPYRIGHT AND LICENSE

Copyright 2002, 2004, 2006--2010, 2012 by Matijs van Zuijlen

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
use strict;
use warnings;
use Carp;
use vars qw($VERSION);
$VERSION = "0.914";
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
      unless ($self->_get_pps_name($child) =~ / ^ ( __recip | __attach
	| __substg1 | __nameid | __properties ) /x
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
	if (eval { $child->isa('OLE::Storage_Lite::PPS::File')}) {
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
    }
  } else {
    $self->_warn_about_unknown_directory($pps);
  }
  return;
}

sub _property_stream_header_length { return 8; }

1;

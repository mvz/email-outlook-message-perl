#!/usr/bin/perl -w
package MSGParser;
use strict;
use OLE::Storage_Lite;
use POSIX qw(mktime ctime);
use constant DIR_TYPE => 1;
use constant FILE_TYPE => 2;

#
# Set up case loop hashes.
#
our $SUBITEM_DIR = {
  '__recip_version1 0_ ' => \&_AddressDir,
  '__attach_version1 0_ ' => \&_AttachmentDir,
};

#
# Main body of module
#
sub new {
  my $that = shift;
  my $class = ref $that || $that;

  my $self = {};
  bless $self, $class;
}

#
# Main sub: parse the PPS tree, and return 
#
sub parse {
  my $self = shift;
  my $file = shift or die "Internal error: no file name";

  # Load and parse MSG file (is OLE)
  my $Msg = OLE::Storage_Lite->new($file);
  my $PPS = $Msg->getPpsTree(1);
  $PPS or die "$file must be an OLE file";

  $self->_RootDir($PPS);
}

sub print {
  my $self = shift;
  print "Date: $self->{OLEDATE}\n" if defined $self->{OLEDATE};
}

#
# Below are functions that walk the PPS tree. The *Dir functions handle
# processing the directory nodes of the tree (mainly, iterating over the
# children), whereas the *Item functions handle processing the items in the
# directory (if such an item is itself a directory, it will in turn be
# processed by the relevant *Dir function).
#

#
# RootItem: Check Root Entry, parse sub-entries.
# The OLE file consists of a single entry called Root Entry, which has
# several children. These children are parsed in the sub SubItem.
# 
sub _RootDir {
  my $self = shift;
  my $PPS = shift;

  my $name = $self->_GetName($PPS);

  $name eq "Root Entry" or warn "Unexpected root entry name $name";

  foreach my $child (@{$PPS->{Child}}) {
    $self->_SubItem($child);
  }
}

sub _SubItem {
  my $self = shift;
  my $PPS = shift;
  
  my $name = $self->_GetName($PPS);
  # DIR Entries:
  if ($PPS->{Type} == DIR_TYPE)
  {
    $self->_GetOLEDate($PPS);
#    if ($name =~ /__recip_version1 0_ /) { # Address of one recipient
#      $self->_AddressDir($PPS);
#    } elsif ($name =~ '__attach_version1 0_ ') { # Attachment
#      $self->_AttachmentDir($PPS);
#    } else {
#      $self->_UnknownEntry($name);
#    }
    my $sub = $SUBITEM_DIR->{$name};
    if (defined $sub) {
      $self->$sub($PPS);
    } else {
      warn "Unknown: $name";
    }
  }
  elsif ($PPS->{Type} == FILE_TYPE)
  {
  }
  else
  {
    warn "Unknown entry type: $PPS->{Type}";
  }
}

sub _AddressDir {
}

sub _AttachmentDir {
}

#
# Helper functions
#

sub _GetName {
  my $self = shift;
  my $PPS = shift;

  return $self->_NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
}

sub _NormalizeWhiteSpace {
  my $self = shift;
  my $name = shift;
  $name =~ s/\W/ /g;
  return $name;
}

sub _GetOLEDate {
  my $self = shift;
  my $PPS = shift;
  unless (defined ($self->{OLEDATE})) {
    # Make Date
    my $date;
    $date = $PPS->{Time2nd};
    $date = $PPS->{Time1st} unless($date);
    if ($date) {
      $self->{OLEDATE} = mktime(@$date);
    }
  }
}

package main;
use Getopt::Long;
use Pod::Usage;

# Setup command line processing.
my $verbose = '';
my $help = '';	    # Print help message and exit.
GetOptions('verbose' => \$verbose, 'help|?' => \$help) or pod2usage(2);
pod2usage(1) if $help;

# Get file name
my $file = $ARGV[0];
defined $file or pod2usage(2);
warn "Will parse file: $file\n" if $verbose; 

# parse file
my $parser = new MSGParser();
$parser->parse($file);
$parser->print();

#
# Usage info follows.
#
__END__

=head1 NAME

msgconvert.pl - Convert Outlook .msg files to mbox format

=head1 SYNOPSIS

msgconvert.pl [options] <file.msg>

  Options:
    --verbose	be verbose
    --help	help message

=head1 OPTIONS

=over 8

=item B<--verbose>

    Print information about skipped parts of the .msg file.

=item B<--help>

    Print a brief help message.

=head1 DESCRIPTION

This program will output the message contained in file.msg in mbox format
on stdout. It will complain about unrecognized OLE parts on
stderr.

=head1 BUGS

Not all data that's in the .MSG file is converted. There simply are some
parts whose meaning escapes me. One of these must contain the date the
message was sent, for example. Formatting of text messages will also be
lost. YMMV.

=cut

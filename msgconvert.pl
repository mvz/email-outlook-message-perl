#!/usr/bin/perl -w
#
# msgconvert.pl:
#
# Convert .MSG files (made by Outlook (Express)) to multipart MIME messages.
#
# Copyright 2002, 2004, 2006, 2007 Matijs van Zuijlen
#
# This program is free software; you can redistribute it and/or modify it
# under the terms of the GNU General Public License as published by the
# Free Software Foundation; either version 2 of the License, or (at your
# option) any later version.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
# Public License for more details.
#

use MSGParser;
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

# parse PPS tree
my $parser = new MSGParser $file, $verbose;
print $parser->as_mbox();

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

#!/usr/bin/perl -w
#
# msgconvert.pl:
#
# Convert .MSG files (made by Outlook Express) to multipart MIME messages.
#
# Copyright 2002 Matijs van Zuijlen
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
# TODO:
# - Wishlist: Make use of more of the items.
# - Make sanity checks when setting mime/mail headers: are duplicates
# created, are values overwritten?
# - Make From (Not From:) line less fake.
# - Refactor handling of more items into separate functions.
# - Wishlist: Parse item names: There must be some structure in them.
#
# CHANGES:
# 20020715  Recognize new items 'Cc', mime type of attachment, long
#	    filename of attachment, and full headers. Attachments turn out
#	    to be numbered, so a regexp is now used to recognize label of
#	    items that are attachments.
# 20020831  long file name will definitely be used if present. Full headers
#	    and mime type information are used when present. Created
#	    generic system for specifying known items to be skipped.
#	    Unexpected contents is never reason to bail out anymore. Added
#	    support for usage message and option processing (--verbose).

#
# Start of main program.
#
use strict;
use Getopt::Long;
use OLE::Storage_Lite;
use MIME::Entity;
use MIME::Parser;
use Pod::Usage;

# Setup command line processing.
my $verbose = '';  # Be verbose about skipped and unknown entries;
my $help = '';	    # Print help message and exit.
GetOptions('verbose' => \$verbose, 'help|?' => \$help) or pod2usage(2);
pod2usage(1) if $help;

# Get file name
my $file = $ARGV[0];
defined $file or pod2usage(2);

# Load and parse MSG file (is OLE)
my $Msg = OLE::Storage_Lite->new($file);
my $PPS = $Msg->getPpsTree(1);
$PPS or die "$file must be an OLE file";

# Create Mime message object
my $Mime = MIME::Entity->build(Type => "multipart/mixed");

# Create and fill Hash to keep lists of entries we skip and their
# (possible) meanings.
my $skippable_entries = { };
FillHashes($skippable_entries);

# We need to keep addresse in a separate structure, because we don't get
# all bits at once.
my $addresses = {   'From' => {},
		    'To' => {},
		    'Reply-To' => {}, };

# Parse the message
RootItem($PPS, $Mime, $addresses, $skippable_entries, $verbose);

# Add the address info to the MIME object.
foreach my $key (keys %$addresses) {
    my $address = $addresses->{$key};
    if (exists $address->{Address}) {
	my $string = "$address->{Name} <$address->{Address}>";
	$string =~ s/\000//g;
	SetHeader($Mime, $key, $string);
    }
}

# print the message to STDOUT using fake From line.
print "From someone\@somewhere Fri Mar 15 00:00:01 1900\n";
$Mime->print(\*STDOUT);
print "\n";

#
# End of main program.
#

#
# RootItem: Check Root Entry, parse sub-entries.
# 
sub RootItem {
    my $PPS = shift;
    my $Mime = shift;
    my $addresses = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));

    $Name eq "Root Entry" or die "Unexpected entry $Name";

    foreach my $Child (@{$PPS->{Child}}) {
	SubItem($Child, $Mime, $addresses, $skippable_entries, $verbose);
    }
}

#
# Parse Level one sub-entries of the Root Entry.
#
sub SubItem {
    my $PPS = shift;
    my $Mime = shift;
    my $addresses = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
    my $Data = $PPS->{Data};

    # DIR Entries:
    if ($PPS->{Type}==1) {
	if ($Name eq '__recip_version1 0_ 00000000') { # Recipient ?
	    ParseAddress($PPS, $addresses, 'To', $skippable_entries,
		$verbose);
	} elsif ($Name eq '__recip_version1 0_ 00000001') { # Sender ?
	    ParseAddress($PPS, $addresses, 'From', $skippable_entries,
		$verbose);
	} elsif ($Name =~ '__attach_version1 0_ ') { # Attachment
	    ParseAttachment($PPS, $Mime, $skippable_entries, $verbose);
	} else {
	    CheckSkippableEntry($Name, "",
		$skippable_entries->{SUB_DIR}, $verbose);
	}
    }
    # FILE Entries.
    elsif ($PPS->{Type}==2) {
	if ($Name eq '__substg1 0_0037001E') {	# Subject
	    SetHeader($Mime, 'Subject', $Data);
	} elsif ($Name eq '__substg1 0_0042001E') {	# Sender Name?
	    SetAddressPart($addresses, 'From', 'Name', $Data, $verbose);
	} elsif ($Name eq '__substg1 0_0065001E') {	# Sender Address?
	    SetAddressPart($addresses, 'From', 'Address', $Data, $verbose);
	} elsif ($Name eq '__substg1 0_0C1A001E') {	# Reply Name?
	    SetAddressPart($addresses, 'Reply-To', 'Name', $Data,
		$verbose);
	} elsif ($Name eq '__substg1 0_0C1F001E') {	# Reply Address?
	    SetAddressPart($addresses, 'Reply-To', 'Address', $Data,
		$verbose);
	} elsif ($Name eq '__substg1 0_0E04001E') {	# Recipient name?
	    SetAddressPart($addresses, 'To', 'Name', $Data, $verbose);
	} elsif ($Name eq '__substg1 0_0E03001E') {	# Cc: Addresses?
	    $Data =~ s/\000//g;
	    unless ($Data eq "") {
		SetHeader($Mime, 'Cc', $Data);
	    }
	} elsif ($Name eq '__substg1 0_1000001E') {	# Body
	    my $handle;
	    my $ent = $Mime->attach(
		    Type => 'text/plain',
		    Encoding => 'quoted-printable',
		    Data => []);
	    $Data =~ s/\015//g;
	    $Data =~ s/\000//g;
	    if ($handle = $ent->open("w")) {
		$handle->print($Data);
		$handle->close;
	    } else {
		warn "Could not write body data!";
	    }
	} elsif ($Name eq '__substg1 0_007D001E') {	# Full headers
	    my $parser = new MIME::Parser;
	    $parser->output_to_core(1);
	    my $entity = $parser->parse_data($Data);
	    my $head = $entity->head;
	    $head->unfold;
	    foreach my $tag ($head->tags) {
		my $writetag;
		if ($tag eq 'Received' or $tag eq 'Date') {
		    $writetag = $tag;
		} else {
		    $writetag = "X-MsgConvert-Original-$tag";
		}
		my @values = $head->get_all($tag);
		foreach my $value (@values) {
		    $Mime->head->add($writetag, $value);
		}
	    };
	} else {
	    CheckSkippableEntry($Name, $Data,
		$skippable_entries->{SUB}, $verbose);
	}
    }
    else {
	warn "Unknown entry type: $PPS->{Type}";
    }
}

#
# Parse an Address type DIR entry.
#
sub ParseAddress {
    my $PPS = shift;
    my $addresses = shift;
    my $entry = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    foreach my $Child (@{$PPS->{Child}}) {
	AddressItem($Child, $addresses, $entry, $skippable_entries,
	    $verbose);
    }
}

#
# Process an item from an address DIR entry.
#
sub AddressItem {
    my $PPS = shift;
    my $addresses = shift;
    my $entry = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
    my $Data = $PPS->{Data};

    # DIR Entries: There should be none.
    if ($PPS->{Type}==1) {
	warn "Unexpected DIR entry $Name";
    }
    # FILE Entries.
    elsif ($PPS->{Type}==2) {
	if ($Name eq '__substg1 0_3001001E') {	# Real Name
	    SetAddressPart($addresses, $entry, 'Name', $Data, $verbose);
	} elsif ($Name eq '__substg1 0_3003001E') {	# Address
	    SetAddressPart($addresses, $entry, 'Address', $Data, $verbose);
	} else {
	    CheckSkippableEntry($Name, $Data,
		$skippable_entries->{ADD}, $verbose)
	}
    }
    else {
	warn "Unknown entry type: $PPS->{Type}";
    }
}

sub SetAddressPart {
    my $addresses = shift;
    my $entry = shift;
    my $part = shift;
    my $data = shift;
    my $verbose = shift;

    if (defined ($addresses->{$entry}->{$part})) {
	if ($addresses->{$entry}->{$part} eq $data) {
	    warn "Skipping duplicate but identical address information for"
		. " $entry/$part\n" if $verbose;
	} else {
	    warn "Address information $entry/$part inconsistent:\n";
	    warn "    Original data: $addresses->{$entry}->{$part}\n";
	    warn "    New data: $data\n";
	}
    } else {
	$addresses->{$entry}->{$part} = $data;
    }
}
#
# Parse an Attachment type DIR entry.
#
sub ParseAttachment {
    my $PPS = shift;
    my $Mime = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    my $ent = $Mime->attach(
	    Type => 'application/octet-stream',
	    Encoding => 'base64',
	    Data => []);
    my $ent_info = {"shortname" => undef,
		    "longname"  => undef };
    foreach my $Child (@{$PPS->{Child}}) {
	AttachItem($Child, $ent, $ent_info, $skippable_entries,
	    $verbose);
    }
    if (defined $ent_info->{"longname"}) {
	$ent->head->mime_attr('content-type.name',
	    $ent_info->{"longname"})
    } elsif (defined $ent_info->{"shortname"}) {
	$ent->head->mime_attr('content-type.name',
	    $ent_info->{"shortname"});
    }
    if (defined $ent_info->{"mimetype"}) {
	$ent->head->mime_attr('content-type',
	    $ent_info->{"mimetype"});
    }
}

#
# Process an item from an attachment DIR entry.
#
sub AttachItem {
    my $PPS = shift;
    my $ent = shift;
    my $ent_info = shift;
    my $skippable_entries = shift;
    my $verbose = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
    my $Data = $PPS->{Data};

    # DIR Entries:
    if ($PPS->{Type}==1) {
	warn "Unknown DIR entry $Name";
    }
    # FILE Entries.
    elsif ($PPS->{Type}==2) {
	if ($Name eq '__substg1 0_37010102') {	# File contents
	    my $handle;
	    if ($handle = $ent->open("w")) {
		$handle->print($Data);
		$handle->close;
	    } else {
		warn "Could not write file data!";
	    }
	} elsif ($Name eq '__substg1 0_3704001E') {	# Short file name
	    $Data =~ s/\000//g;
	    $ent_info->{"shortname"} = $Data;
	} elsif ($Name eq '__substg1 0_3707001E') {	# Long file name
	    $Data =~ s/\000//g;
	    $ent_info->{"longname"} = $Data;
	} elsif ($Name eq '__substg1 0_370E001E') {	# mime type
	    $Data =~ s/\000//g;
	    $ent_info->{"mimetype"} = $Data;
	} else {
	    CheckSkippableEntry($Name, $Data,
		$skippable_entries->{ATT}, $verbose)
	}
    }
    else {
	warn "Unknown entry type: $PPS->{Type}";
    }
}

#
# Set a mime header, taking care of cleaning up the data.
#
sub SetHeader {
    my $Mime = shift;
    my $Item = shift;
    my $Value = shift;

    $Value =~ s/\000//g;

    $Mime->head->replace($Item, $Value);
}

#
# Replace whitespace (typically, in an entry name) with a single space.
#
sub NormalizeWhiteSpace {
    my $Name = shift;
    $Name =~ s/\W/ /g;
    return $Name;
}

sub FillHashes {
    my $skippable_entries = shift;

    $skippable_entries->{SUB_DIR} = {
	'__nameid_version1 0' => ["Introductory stuff"]
    };
    $skippable_entries->{SUB} = {
	'__substg1 0_001A001E' => ["Type of message", "IPM.Note"],
	'__substg1 0_003B0102' => ["Sender address variant"],
	'__substg1 0_003D001E' => ["Contains 'Re: '", "Re: "],
	'__substg1 0_00410102' => ["Sender address variant"],
	'__substg1 0_0064001E' => ["Sender address type", "SMTP"],
	'__substg1 0_0070001E' => ["Subject w/o Re"],
	'__substg1 0_00710102' => ["16 bytes: Unknown"],
	'__substg1 0_0C190102' => ["Reply address variant"],
	'__substg1 0_0C1D0102' => ["Reply address variant"],
	'__substg1 0_0C1E001E' => ["Reply address type", "SMTP"],
	'__substg1 0_0E02001E' => ["1 byte: Unknown", ""],
	'__substg1 0_0E1D001E' => ["Subject w/o Re"],
	'__substg1 0_1008001E' => ["Summary or something"],
	'__substg1 0_10090102' => ["Binary data, may be largish"],
	'__substg1 0_300B0102' => ["16 bytes: Unknown"],
	'__substg1 0_3FF8001E' => ["Routing data"],
	'__substg1 0_3FF90102' => ["Routing data"],
	'__substg1 0_3FFA001E' => ["Routing data"],
	'__substg1 0_3FFB0102' => ["Routing data"],
	'__properties_version1 0' => ["Properties"],
    };
    $skippable_entries->{ADD} = {
	'__properties_version1 0' => ["Properties"],
	'__substg1 0_0FF60102' => ["Index"],
	'__substg1 0_0FFF0102' => ["Address variant"],
	'__substg1 0_3002001E' => ["Address Type", "SMTP"],
	'__substg1 0_300B0102' => ["Address variant"],
	'__substg1 0_3A20001E' => ["Address variant"],
    };
    $skippable_entries->{ATT} = {
	'__substg1 0_0FF90102' => ["Index"],
	'__properties_version1 0' => ["Properties"],
    };
}

sub CheckSkippableEntry {
    my $name = shift;
    my $data = shift;
    my $entries = shift;
    my $verbose = shift;

    $data =~ s/\000//g;

    my $msg = "Skipping entry $name : ";
    if (exists $entries->{$name}) {
	my $entrydata = $entries->{$name};
	$msg .= "$entrydata->[0]";
	if (defined ($entrydata->[1])) {
	    unless ($data eq $entrydata->[1]) {
		$msg .= " [UNEXPECTED DATA]";
	    }
	}
	warn "$msg\n" if $verbose;
    } else {
	$msg .= "UNKNOWN";
	warn "$msg\n";
    }
}

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

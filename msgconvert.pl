#!/usr/bin/perl -w
#
# msgconvert.pl:
#
# Convert .MSG files (made by Outlook (Express)) to multipart MIME messages.
#
# Copyright 2002, 2004 Matijs van Zuijlen
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
# TODO (Functional)
# - Make sanity checks when setting mime/mail headers: are duplicates
#   created, are values overwritten?
# - Use address data to make To: and Cc: lines complete
# - Fix full header parsing, and use it to get more data
# - Make use of more of the items, if possible.
# TODO (Technical)
# - Do not use Mime object for storage while parsing the file. This will
#   make it possible to postpone certain decisions until all all the facts
#   are known.
# - Refactor handling of more items into separate functions.
# - Parse item names: There must be some structure in them.
# - Make functional part into a module.
#
# CHANGES:
# 20040214  Fix typos and incorrect comments.
# 20040104  Handle address data slightly better, make From line less fake,
#	    make $verbose and $skippable_entries global vars, handle HTML
#	    variant of body text if present (though not optimally).
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
# Import modules.
#
use strict;
use Getopt::Long;
use OLE::Storage_Lite;
use MIME::Entity;
use MIME::Parser;
use Pod::Usage;
use POSIX qw(mktime ctime);

#
# Note: These are the only two global variables. They're global since
# they're almost like constants.
#
my $verbose;		# Be verbose about skipped and unknown entries;
my $skippable_entries;	# Entries we don't want to use

main();

#
# Subroutines go below.
# 
sub main {
    # Setup command line processing.
    $verbose = '';
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

    # Fill Hash to keep lists of entries we skip and their (possible)
    # meanings.
    $skippable_entries = {};
    FillHashes($skippable_entries);

    # We need to keep addresses in a separate structure, because we don't get
    # all bits at once.
    my $addresses = {
	'From' => {},
	'Reply-To' => {},
	'ToList' => {},
	'CcList' => {},
    };

    # Parse the message
    RootItem($PPS, $Mime, $addresses);

    # Add the address info to the MIME object.
    foreach my $key (keys %$addresses) {
	my $address = $addresses->{$key};
	AddAddressHeader($Mime, $key, $address);
    }

    # Construct From line from whatever we know.
    my $string = "";
    if (exists $addresses->{'From'}->{'Address'}) {
	$string = $addresses->{'From'}->{'Address'};
    }
    if($string =~ /@/) {
    	$string =~ s/\n//g;
    } else {
	$string = 'someone@somewhere';
    }
    print "From ", $string, " ", $Mime->head->get('Date') ;
    $Mime->print(\*STDOUT);
    print "\n";
}

#
# RootItem: Check Root Entry, parse sub-entries.
# The OLE file consists of a single entry called Root Entry, which has
# several children. These children are parsed in the sub SubItem.
# 
sub RootItem {
    my $PPS = shift;
    my $Mime = shift;
    my $addresses = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));

    $Name eq "Root Entry" or warn "Unexpected root entry name $Name";

    foreach my $Child (@{$PPS->{Child}}) {
	SubItem($Child, $Mime, $addresses);
    }
}

#
# Parse Level one sub-entries of the Root Entry.
# Thes can be DIR or FILE type entries.
#
sub SubItem {
    my $PPS = shift;
    my $Mime = shift;
    my $addresses = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
    my $Data = $PPS->{Data};

    # DIR Entries:
    if ($PPS->{Type}==1) {
	# Set the OLE date, if we haven't found a date in another way.
	if (! $Mime->head->count('Date')) {
	    # Make Date
	    my $date;
	    $date = $PPS->{Time2nd};
	    $date = $PPS->{Time1st} unless($date);
	    if ($date) {
		my $time = mktime(@$date);
		my $datestring = sprintf(
		    "%02d.%02d.%4d %02d:%02d:%02d",
		    $date->[3], $date->[4]+1, $date->[5]+1900,
		    $date->[2], $date->[1],   $date->[0]
		);
		SetHeader($Mime, 'Date', ctime($time));
	    }
	}
	if ($Name =~ /__recip_version1 0_ /) { # Address of one recipient
	    ParseAddress($PPS, $Mime, 'Addresses');
	} elsif ($Name =~ '__attach_version1 0_ ') { # Attachment
	    ParseAttachment($PPS, $Mime);
	} else {
	    CheckSkippableEntry($Name, "", "SUB_DIR");
	}
    }
    # FILE Entries.
    elsif ($PPS->{Type}==2) {
	if ($Name eq '__substg1 0_0037001E') {	# Subject
	    SetHeader($Mime, 'Subject', $Data);
	} elsif ($Name eq '__substg1 0_0042001E') {	# From: Name?
	    SetAddressPart($addresses->{'From'}, 'Name', $Data);
	} elsif ($Name eq '__substg1 0_0065001E') {	# From: Address?
	    SetAddressPart($addresses->{'From'}, 'Address', $Data);
	} elsif ($Name eq '__substg1 0_0C1A001E') {	# Reply-To: Name?
	    SetAddressPart($addresses->{'Reply-To'}, 'Name', $Data);
	} elsif ($Name eq '__substg1 0_0C1F001E') {	# Reply-To: Address?
	    SetAddressPart($addresses->{'Reply-To'}, 'Address', $Data);
	} elsif ($Name eq '__substg1 0_0E04001E') {	# To: Names
	    SetAddressPart($addresses->{'ToList'}, 'Name', $Data);
	} elsif ($Name eq '__substg1 0_0E03001E') {	# Cc: Names
	    SetAddressPart($addresses->{'CcList'}, 'Name', $Data);
	} elsif ($Name eq '__substg1 0_1000001E') {	# Body
	    SaveBody($Mime, $Data, 'text/plain; charset=ISO-8859-1');
	} elsif ($Name eq '__substg1 0_10130102') {	# HTML Version of body
	    SaveBody($Mime, $Data, "text/html");
	} elsif ($Name eq '__substg1 0_007D001E') {	# Full headers
	    warn "Adding full header info:\n";
	    my $parser = new MIME::Parser;
	    $parser->output_to_core(1);
	    $parser->decode_headers(1);
	    #$parser->ignore_errors(0);
	    my $entity = $parser->parse_data($Data)
		or warn "Couldn't parse full headers!"; 
	    # $entity->dump_skeleton;
	    my $head = $entity->head;
	    $head->unfold;
	    foreach my $tag ($head->tags) {
		my $writetag;
		if ($tag eq 'Received' or $tag eq 'Date') {
		    $writetag = $tag;
		} else {
		    $writetag = "X-MsgConvert-Original-$tag";
		}
		warn "Tag: $tag -> $writetag\n";
		my @values = $head->get_all($tag);
		foreach my $value (@values) {
		    $Mime->head->add($writetag, $value);
		}
	    };
	} else {
	    CheckSkippableEntry($Name, $Data, "SUB");
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
    my $Mime = shift;
    my $entry = shift;

    #warn "Processing address subitem $entry\n";
    my $address = {};
    foreach my $Child (@{$PPS->{Child}}) {
	AddressItem($Child, $address, $entry);
    }
    AddAddressHeader($Mime, $entry, $address);
}

#
# Process an item from an address DIR entry.
#
sub AddressItem {
    my $PPS = shift;
    my $address = shift;

    my $Name = NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($PPS->{Name}));
    my $Data = $PPS->{Data};

    # DIR Entries: There should be none.
    if ($PPS->{Type}==1) {
	warn "Unexpected DIR entry $Name";
    }
    # FILE Entries.
    elsif ($PPS->{Type}==2) {
	if ($Name eq '__substg1 0_3001001E') {	# Real Name
	    SetAddressPart($address, 'Name', $Data);
	} elsif ($Name eq '__substg1 0_3003001E') {	# Address
	    SetAddressPart($address, 'Address', $Data) if ($Data =~ /@/);
	} elsif ($Name eq '__substg1 0_403E001E') {	# Address
	    SetAddressPart($address, 'Address', $Data) if ($Data =~ /@/);
	} else {
	    CheckSkippableEntry($Name, $Data, "ADD")
	}
    }
    else {
	warn "Unknown entry type: $PPS->{Type}";
    }
}

sub SetAddressPart {
    my $address = shift;
    my $partname = shift;
    my $data = shift;

    $data =~ s/\000//g;
    #warn "Processing address data part $partname : $data\n";
    if (defined ($address->{$partname})) {
	if ($address->{$partname} eq $data) {
	    warn "Skipping duplicate but identical address information for"
	    . " $partname\n" if $verbose;
	} else {
	    warn "Address information $partname inconsistent:\n";
	    warn "    Original data: $address->{$partname}\n";
	    warn "    New data: $data\n";
	}
    } else {
	$address->{$partname} = $data;
    }
}

#
# Parse an Attachment type DIR entry.
#
sub ParseAttachment {
    my $PPS = shift;
    my $Mime = shift;

    my $ent = $Mime->attach(
	Type => 'application/octet-stream',
	Encoding => 'base64',
	Data => []
    );
    my $ent_info = {
	"shortname" => undef,
	"longname"  => undef
    };
    foreach my $Child (@{$PPS->{Child}}) {
	AttachItem($Child, $ent, $ent_info);
    }
    if (defined $ent_info->{"longname"}) {
	$ent->head->mime_attr('content-type.name', $ent_info->{"longname"})
    } elsif (defined $ent_info->{"shortname"}) {
	$ent->head->mime_attr('content-type.name', $ent_info->{"shortname"});
    }
    if (defined $ent_info->{"mimetype"}) {
	$ent->head->mime_attr('content-type', $ent_info->{"mimetype"});
    }
}

#
# Process an item from an attachment DIR entry.
#
sub AttachItem {
    my $PPS = shift;
    my $ent = shift;
    my $ent_info = shift;

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
	    CheckSkippableEntry($Name, $Data, "ATT");
	}
    }
    else {
	warn "Unknown entry type: $PPS->{Type}";
    }
}

sub SaveBody {
    my $Mime = shift;
    my $Data = shift;
    my $type = shift;

    my $handle;
    my $ent = $Mime->attach(
	# Type => 'text/plain',
	Type => $type,
	# Encoding => 'quoted-printable',
	Encoding => '8bit',
	Data => []
    );
    $Data =~ s/\015//g;
    $Data =~ s/\000//g;
    if ($handle = $ent->open("w")) {
	$handle->print($Data);
	$handle->close;
    } else {
	warn "Could not write body data!";
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

    # $Mime->head->replace($Item, $Value);
    $Mime->head->add($Item, $Value);
}

#
# Add an email address to the mime header, given a structure:
#
sub AddAddressHeader {
    my $Mime = shift;
    my $key = shift;
    my $address = shift;

    my $string = "$address->{Name}";
    if (not exists $address->{Address}) {
	$string .= "";
    } elsif ($address->{Address} =~ /@/) {
	$string .= " <$address->{Address}>";
    } else {
	$string .= " <no_valid_address>";
    }
    unless ($string eq "") {
	$string =~ s/\000//g;
	SetHeader($Mime, $key, $string);
    }
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
	'__substg1 0_00460102' => ["Sender address variant"],
	'__substg1 0_00530102' => ["Sender address variant"],
	'__substg1 0_0064001E' => ["Sender address type", "SMTP"],
	'__substg1 0_0070001E' => ["Subject w/o Re"],
	'__substg1 0_00710102' => ["16 bytes: Unknown"],
	'__substg1 0_0C190102' => ["Reply address variant"],
	'__substg1 0_0C1D0102' => ["Reply address variant"],
	'__substg1 0_0C1E001E' => ["Reply address type", "SMTP"],
	'__substg1 0_0E02001E' => ["1 byte: Unknown", ""],
	'__substg1 0_0E1D001E' => ["Subject w/o Re"],
	'__substg1 0_0E270102' => ["64 bytes: Unknown"],
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
	'__substg1 0_0FF60102' => ["Index"],
	'__substg1 0_0FFF0102' => ["Address variant"],
	'__substg1 0_3002001E' => ["Address Type", "SMTP"],
	'__substg1 0_300B0102' => ["Address variant"],
	'__substg1 0_3A20001E' => ["Address variant"],
	'__properties_version1 0' => ["Properties"],
    };
    $skippable_entries->{ATT} = {
	'__substg1 0_0FF90102' => ["Index"],
	'__properties_version1 0' => ["Properties"],
    };
}

sub CheckSkippableEntry {
    my $name = shift;
    my $data = shift;
    my $subset = shift;

    my $entries = $skippable_entries->{$subset};

    $data =~ s/\x00//g;
    $data =~ s/[\x01-\x1F]/_/g;
    $data =~ s/[\x7F-\xFF]/_/g;

    my $msg = "Skipping entry $name : ";
    $msg .= substr($data,0,20);
    if (exists $entries->{$name}) {
	my $entrydata = $entries->{$name};
	$msg .= " ($entrydata->[0]) ";
	if (defined ($entrydata->[1])) {
	    unless ($data eq $entrydata->[1]) {
		$msg .= " [UNEXPECTED VALUE]";
	    }
	}
	warn "$msg\n" if $verbose;
    } else {
	$msg .= " (UNKNOWN) ";
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

#!/usr/bin/perl -w
# oledump.pl
#
# Based on:
#
# OLE::Storage_Lite Sample
# Name : smplls.pl
#  by Kawai, Takanori (Hippo2000) 2000.11.4
# Displays PPS structure of specified file
# Just subset of lls that is distributed with OLE::Storage
#
# Changes made by Matijs van Zuijlen:
#   2002.04.16:
#	Indenting
#	Added #! line.
#	Moved main of program to end.
#	English
#	Removed prototype info from PrnItem, awaiting knowledge on how to
#	resolve warnings about it.
#	Added code to print data as well.
#   2002.04.17:
#	Allow more characters as printable.
#
# Usage:
#   oledump.pl <file> [print_data]
#
#   will show the OLE structure of <file>. setting print_data to 1 will
#   also dump the data parts, showing hex and any printable characters.
#
#=================================================================

use strict;
use OLE::Storage_Lite;
use locale;

#----------------------------------------------------------------
# PrnItem: Displays PPS infomations
#----------------------------------------------------------------
sub PrnItem {
    my($oPps, $iLvl, $iTtl, $iDir, $prData) = @_;
    my $raDate;
    my %sPpsName = (1 => 'DIR', 2 => 'FILE', 5=>'ROOT');

    # Make Name (including PPS-no and level)
    my $sName = OLE::Storage_Lite::Ucs2Asc($oPps->{Name});
    $sName =~ s/\W/ /g;
    $sName = sprintf("%s %3d '%s' (pps %x)", 
	    ' ' x ($iLvl * 2), $iDir, $sName, $oPps->{No});

    # Make Date 
    my $sDate;
    if($oPps->{Type}==2) {
	$sDate = sprintf("%10x bytes", $oPps->{Size});
    }
    else {
	$raDate = $oPps->{Time2nd};
	$raDate = $oPps->{Time1st} unless($raDate);
	$sDate = ($raDate)?
	    sprintf("%02d.%02d.%4d %02d:%02d:%02d", 
		    $raDate->[3], $raDate->[4]+1, $raDate->[5]+1900,
		    $raDate->[2], $raDate->[1],   $raDate->[0]) : "";
    }

    # Display
    printf "%02d %-50s %-4s %s\n", 
	${$iTtl}++,
	$sName,
	$sPpsName{$oPps->{Type}},
	$sDate;
    
    # MvZ: Print Data
    if ($prData and $oPps->{Type}==2 and $oPps->{Size} > 0) {
	my $data = $oPps->{Data};
	my $length = length($data);
	my $numloops = $length/16;
	my $i;

	for ($i=0; $i<$numloops; $i++) {
	    #print "$i; $numloops;\n";
	    my $substring = substr($data, $i*16, 16);
	    my $copy = $substring;
	    $substring =~ s/./sprintf("%02x ", ord($&))/sge;
	    $copy =~ s/[^[:print:]]/./sg;
	    #$copy =~ s/[\x00-\x1f\x7f-\xa0]/./sg;
	    #$copy =~ s/[\x09\x0a\x0c\x0d]/./sg;

	    print " " x 12;
	    print sprintf("%-48s %-16s\n", $substring, $copy);
	}
    }

    # For its Children
    my $iDirN=1;
    foreach my $iItem (@{$oPps->{Child}}) {
	PrnItem($iItem, $iLvl+1, $iTtl, $iDirN, $prData);
	$iDirN++;
    }
}

# Main
die "No files are specified" if($#ARGV < 0);
my $oOl = OLE::Storage_Lite->new($ARGV[0]);
my $prData = $ARGV[1];
my $oPps = $oOl->getPpsTree(1);
die( $ARGV[0]. " must be a OLE file") unless($oPps);
my $iTtl = 0;
PrnItem($oPps, 0, \$iTtl, 1, $prData);

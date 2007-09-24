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
#=================================================================

use strict;
use OLE::Storage_Lite;
use Getopt::Long;
use Pod::Usage;
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
#
my $prData;
my $help = '';	    # Print help message and exit.
my $opt = GetOptions("with-data" => \$prData) or pod2usage(2);
pod2usage(1) if $help;
pod2usage(2) if($#ARGV < 0);
foreach my $file (@ARGV) {
  my $oOl = OLE::Storage_Lite->new($file);
  my $oPps = $oOl->getPpsTree(1);
  die( $file. " must be a OLE file") unless($oPps);
  my $iTtl = 0;
  PrnItem($oPps, 0, \$iTtl, 1, $prData);
}
#
# Usage info follows.
#
__END__

=head1 NAME

oledump.pl - Dump structure of an OLE file.

=head1 SYNOPSIS

oledump.pl [options] <file>...

  Options:
    --with-data	    dump data too
    --help	    help message

=head1 OPTIONS

=over 8

=item B<--with-data>

    Dump data as will, showing both hex and any printable characters.

=item B<--help>

    Print a brief help message.

=head1 DESCRIPTION

This program will dump the PPS structure of OLE files passed to it on the
command line. It is based on smplls.pl by Kawai, Takanori, which is part of
the OLE::Storage_Lite distribution.

=cut

#!/usr/bin/perl
use strict;
use warnings;
use Test::Simple tests => 2;
use OLE::Storage_Lite;
use MSGParser;

my $p = new MSGParser('t/files/gpg_signed.msg');
ok($p);
my $m = $p->mime_object;
ok($m->mime_type eq 'multipart/signed');

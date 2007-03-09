#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 3;
use MSGParser;
use MIME::Entity;
use Email::MIME::Creator;

my $p = MSGParser->_empty_new();
ok($p, 'Checking internal new');

my $mime = Email::MIME->create(body => "Hello!");
$p->{HEAD} = <<HEADER;
From: quux\@zonk
MIME-Version: ignore
Content-Type: ignore
Content-Transfer-Encoding: ignore
X-Mailer: ignore
X-Msgconvert: ignore
X-MS-TNEF-Correlator: ignore_case
X-MS-Has-Attach: ignore
HEADER
my @expected_tags = qw{
  Date
  From
  MIME-Version
  };
$p->_copy_header_data($mime);
my @new_tags = $mime->header_names;

is_deeply([sort @new_tags], [sort @expected_tags],
  'Are the right headers inserted?');
isnt($mime->header('MIME-Version'), 'ignore'); 

#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 3;
use MSGParser;
use MIME::Entity;

my $p = MSGParser->_empty_new();
ok($p, 'Checking internal new');

my $mime = MIME::Entity->build(Data => "Hello!");
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
  From
  Content-Type
  Content-Disposition
  Content-Transfer-Encoding
  MIME-Version
  X-Mailer
  };
$p->_copy_header_data($mime);
my $head = $mime->head;
my @new_tags = $head->tags;

is_deeply([sort @new_tags], [sort @expected_tags],
  'Are the right headers inserted?');
isnt($head->get('MIME-Version'), 'ignore'); 



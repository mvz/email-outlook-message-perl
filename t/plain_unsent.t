#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 4;
use MSGParser;
use Email::MIME;

my $p = new MSGParser('t/files/plain_unsent.msg');
ok($p, "Parsing succeeded");
my $m = Email::MIME->new($p->as_string);
like($m->content_type, qr{^text/plain}, "Content type should be text/plain");
is($m->header("Subject"), "Test for MSGConvert -- plain text", "Testing subject");
is($m->header("Date"), "Mon, 26 Feb 2007 22:56:40 +0100", "Testing date");
# TODO: Is this the behavior we want?
is($m->header("From"), "<>", "Testing from");
like($m->header("To"), 'someone@somewhere.com', "Testing to");

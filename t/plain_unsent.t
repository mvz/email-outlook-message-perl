#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 6;
use Email::Outlook::MSG;

my $p = new Email::Outlook::MSG('t/files/plain_unsent.msg');
ok($p, "Parsing succeeded");
my $m = $p->to_email_mime;
like($m->content_type, qr{^text/plain}, "Content type should be text/plain");
is($m->header("Subject"), "Test for MSGConvert -- plain text", "Testing subject");
is($m->header("Date"), "Mon, 26 Feb 2007 22:56:40 +0100", "Testing date");
is($m->header("From"), undef, "No from specified");
like($m->header("To"), qr{someone\@somewhere\.com}, "Testing to");

#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 9;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/plain_unsent.msg');
ok($p, "Parsing succeeded");
my $m = $p->to_email_mime;
is(scalar($m->header_names), 7, "Seven headers");
like($m->content_type, qr{^text/plain}, "Content type should be text/plain");
is($m->header("Subject"), "Test for MSGConvert -- plain text", "Testing subject");
is($m->header("Date"), "Mon, 26 Feb 2007 22:56:40 +0000", "Testing date");
is($m->header("From"), undef, "No from specified");
is($m->header("To"), "Someone Else <someone\@somewhere\.com>", "Testing to");
is($m->body, "This is a test\nThe body is in plain text", "Check body");
is(scalar($m->subparts), 0, "No sub-parts"); 

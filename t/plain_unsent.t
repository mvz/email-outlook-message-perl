# Test unsent plain text message, single-byte encoded.
use strict;
use warnings;
use Test::More tests => 15;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/plain_unsent.msg');
ok($p, "Parsing succeeded");
like($p->get_mapi_property('1000')->[1], qr/This is a test..The body is in plain text/s);

my $m = $p->to_email_mime;
is(scalar($m->header_names), 6, "Six headers");

is($m->header("Subject"), "Test for MSGConvert -- plain text", "Testing subject");
is($m->header("Date"), "Mon, 26 Feb 2007 22:56:40 +0000", "Testing date");
is($m->header("From"), undef, "No from specified");
is($m->header("To"), "Someone Else <someone\@somewhere\.com>", "Testing to");

like($m->content_type, qr{^multipart/alternative}, "Content type should be multipart/alternative");
my @parts = $m->subparts;
is(scalar(@parts), 2, "Two sub-parts");

my $text = $parts[0];
like($text->content_type, qr{^text/plain}, "Content type should be text/plain");
is($text->body, "This is a test\r\nThe body is in plain text", "Check body");
is(scalar($text->subparts), 0, "No sub-parts");

my $rtf = $parts[1];
like($rtf->content_type, qr{^application/rtf}, "Content type should be application/rtf");
is($rtf->header("Content-Disposition"), "inline", "Testing content disposition");
is(scalar($rtf->subparts), 0, "No sub-parts");

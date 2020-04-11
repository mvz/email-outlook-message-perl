# Test charset handling.
use strict;
use warnings;
use Test::More tests => 18;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/charset.msg');
ok($p, "Parsing succeeded");
my $m = $p->to_email_mime;
is(scalar($m->header_names), 43, "43 headers");
like($m->content_type, qr{^multipart/alternative}, "Content type should be multipart/alternative");
is($m->header("Subject"), "PST Export - Embedded Email Test", "Testing subject");
is($m->header("Date"), "Wed, 9 Oct 2019 05:55:10 +0000", "Testing date");
is($m->header("From"), "Joseph Q Bloggs <joebloggs\@example.org>", "From header");
is($m->header("To"), "Joseph Q Bloggs <joebloggs\@example.org>", "Testing to");
is($m->body, "\r\n", "No body");

my @parts = $m->subparts;
is(scalar(@parts), 2, "Two sub-parts");

my $text = $parts[0];
like($text->content_type, qr{^text/plain}, "Content type should be multipart/alternative");
like($text->content_type, qr{; charset="CP1252"}, "charset should be CP1252");
is($text->header("Content-Disposition"), "inline", "Testing content disposition");
is($text->body, "This email contains an email\x85 Email-ception!!!\n\n", "Testing body");
is(scalar($text->subparts), 0, "No sub-parts");
my $html = $parts[1];
like($html->content_type, qr{^text/html}, "Content type should be text/html");
like($text->content_type, qr{; charset="CP1252"}, "charset should be CP1252");
is($html->header("Content-Disposition"), "inline", "Testing content disposition");
is(scalar($html->subparts), 0, "No sub-parts");

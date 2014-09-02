# Test plain text message, jpeg attached.
use strict;
use warnings;
use Test::More tests => 23;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/plain_jpeg_attached.msg');
ok($p, "Parsing succeeded");
my $m = $p->to_email_mime;
is(scalar($m->header_names), 14, "Fourteen headers");
like($m->content_type, qr{^multipart/mixed}, "Content type should be multipart/mixed");
is($m->header("Content-Disposition"), "inline", "Testing content disposition");
is($m->header("Subject"), "test", "Testing subject");
is($m->header("Date"), "Mon, 24 Sep 2007 15:28:03 +0200", "Testing date");
is($m->header("From"), "Matijs van Zuijlen <Matijs.van.Zuijlen\@xs4all.nl>", "From header");
is($m->header("To"), "matijs\@xxxxxx.nl", "Testing to");
is($m->body, "\r\n", "No body");

my @parts = $m->subparts;
is(scalar(@parts), 2, "Two sub-parts");

my $text = $parts[0];
like($text->content_type, qr{^multipart/alternative}, "Content type should be multipart/alternative");
my @textparts = $text->subparts;
is(scalar(@textparts), 2, "Two alternatives");
my $plain = $textparts[0];
like($plain->content_type, qr{^text/plain}, "Content type should be text/plain");
is($plain->header("Content-Disposition"), "inline", "Testing content disposition");
is($plain->body, "test\r\n", "Testing body");
is(scalar($plain->subparts), 0, "No sub-parts");
my $rtf = $textparts[1];
like($rtf->content_type, qr{^application/rtf}, "Content type should be application/rtf");
is($rtf->header("Content-Disposition"), "inline", "Testing content disposition");
is(scalar($rtf->subparts), 0, "No sub-parts");

my $jpg = $parts[1];
like($jpg->content_type, qr{^image/jpeg}, "Content type should be image/jpeg");
is($jpg->header("Content-Disposition"), "attachment; filename=\"test.jpg\"", "Testing content disposition");
is(scalar($jpg->subparts), 0, "No sub-parts");
is(length($jpg->body), 7681, "Correct jpeg size");

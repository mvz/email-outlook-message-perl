# Test unsent plain text message, unicode encoded.
use strict;
use warnings;
use Test::More tests => 15;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/plain_uc_unsent.msg');
ok($p, "Parsing succeeded");
my $m = $p->to_email_mime;
is(scalar($m->header_names), 6, "Six headers");

is($m->header("Subject"), "Test for MSGConvert -- plain text", "Testing subject");
is($m->header("Date"), "Mon, 26 Feb 2007 22:57:01 +0000", "Testing date");
is($m->header("From"), undef, "No from specified");
is($m->header("To"), "Someone Else <someone\@somewhere\.com>", "Testing to");

like($m->content_type, qr{^multipart/alternative}, "Content type should be multipart/alternative");
my @parts = $m->subparts;
is(scalar(@parts), 2, "Two sub-parts"); 

my $text = $parts[0];
like($text->content_type, qr{^text/plain}, "Content type should be text/plain");
is($text->body, "This is a test\nThe body is in plain text\n", "Check body");
is(scalar($text->subparts), 0, "No sub-parts"); 

my $rtf = $parts[1];
like($rtf->content_type, qr{^application/rtf}, "Content type should be application/rtf");
is($rtf->header("Content-Disposition"), "inline", "Testing content disposition");
is($rtf->body, "{\\rtf1\\ansi\\ansicpg1252\\fromtext \\deff0{\\fonttbl\n\r"
  . "{\\f0\\fswiss Arial;}\n\r"
  . "{\\f1\\fmodern Courier New;}\n\r"
  . "{\\f2\\fnil\\fcharset2 Symbol;}\n\r"
  . "{\\f3\\fmodern\\fcharset0 Courier New;}}\n\r"
  . "{\\colortbl\\red0\\green0\\blue0;\\red0\\green0\\blue255;}\n\r"
  . "\\uc1\\pard\\plain\\deftab360 \\f0\\fs20 This is a test\\par\n\r"
  . "The body is in plain text\\par\n\r"
  . "}", "Check RTF");
is(scalar($rtf->subparts), 0, "No sub-parts"); 

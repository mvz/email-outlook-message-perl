#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 2;
use Email::Outlook::MSG;

my $p = new Email::Outlook::MSG('t/files/gpg_signed.msg');
ok($p, "Parsing succeeded");
TODO: {
  local $TODO = "GPG Parsing doesn't work yet";
  my $m = $p->to_email_mime;
  like($m->content_type, qr{^multipart/signed},
    "Content type should be multipart/signed");
}

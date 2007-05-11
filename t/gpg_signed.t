#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 2;
use Email::MSG;
use Email::MIME::Modifier;

my $p = new Email::MSG('t/files/gpg_signed.msg');
ok($p, "Parsing succeeded");
TODO: {
  local $TODO = "GPG Parsing doesn't work yet";
  my $m = Email::MIME->new($p->as_string);
  like($m->content_type, qr{^multipart/signed},
    "Content type should be multipart/signed");
}

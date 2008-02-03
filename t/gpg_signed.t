use strict;
use warnings;
use Test::More tests => 2;
use Email::Outlook::Message;

my $p = new Email::Outlook::Message('t/files/gpg_signed.msg');
ok($p, "Parsing succeeded");
TODO: {
  local $TODO = "GPG Parsing doesn't work yet";
  my $m = $p->to_email_mime;
  like($m->content_type, qr{^multipart/signed},
    "Content type should be multipart/signed");
}

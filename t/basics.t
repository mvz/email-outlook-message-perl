use strict;
use warnings;
use Test::More tests => 1;
use Email::Outlook::Message;

eval {
  my $p = Email::Outlook::Message->new();
};
like($@, qr/^File name is required/);


use strict;
use warnings;
use Test::More tests => 18;
use Email::Outlook::Message;
#use MIME::Entity;
use Email::MIME::Creator;

my $p = Email::Outlook::Message->_empty_new();
ok($p, 'Checking internal new');
test_copy_header_data($p);
test_is_transmittable_property($p);
test_submission_id_date($p);

# DONE

sub test_copy_header_data {
  my $p = shift;

  my $mime = Email::MIME->create(body => "Hello!");
  $p->{HEAD} = <<HEADER;
From: quux\@zonk
MIME-Version: ignore
Content-Type: ignore
Content-Transfer-Encoding: ignore
X-Mailer: ignore
X-Msgconvert: ignore
X-MS-TNEF-Correlator: ignore_case
X-MS-Has-Attach: ignore
HEADER
  my @expected_tags = qw{
  Date
  From
  MIME-Version
  };
  $p->_copy_header_data($mime);
  my @new_tags = $mime->header_names;
  is_deeply([sort @new_tags], [sort @expected_tags],
    'Are the right headers inserted?');
  isnt($mime->header('MIME-Version'), 'ignore'); 
}

sub test_is_transmittable_property {
  my $p = shift;
  ok($p->_is_transmittable_property('0000'));
  ok($p->_is_transmittable_property('0DFF'));
  ok(not $p->_is_transmittable_property('0E00'));
  ok(not $p->_is_transmittable_property('0FFF'));
  ok($p->_is_transmittable_property('1000'));
  ok($p->_is_transmittable_property('5FFF'));
  ok(not $p->_is_transmittable_property('6000'));
  ok(not $p->_is_transmittable_property('67FF'));
  ok($p->_is_transmittable_property('6800'));
  ok($p->_is_transmittable_property('7BFF'));
  ok(not $p->_is_transmittable_property('7C00'));
  ok(not $p->_is_transmittable_property('7FFF'));
  ok($p->_is_transmittable_property('8000'));
  ok($p->_is_transmittable_property('FFFF'));
}

sub test_submission_id_date {
  my $p = shift;
  $p->{SUBMISSION_ID} = "c=us;a=;p=Something;l=ABCDEFGH1-030728080154Z-268.";
  is($p->_submission_id_date, "Mon, 28 Jul 2003 08:01:54 +0000");
}


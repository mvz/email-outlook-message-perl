use strict;
use warnings;
use Test::More tests => 17;
use Email::Outlook::Message;
#use MIME::Entity;
use Email::MIME::Creator;

my $p = Email::Outlook::Message->_empty_new();
ok($p, 'Checking internal new');
test_copy_header_data($p);
test_submission_id_date($p);
test_to_email_mime_with_no_parts($p);
test_to_email_mime_with_plain_part($p);
test_to_email_mime_with_html_part($p);
test_to_email_mime_with_two_parts($p);

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

sub test_submission_id_date {
  my $p = shift;
  $p->{SUBMISSION_ID} = "c=us;a=;p=Something;l=ABCDEFGH1-030728080154Z-268.";
  is($p->_submission_id_date, "Mon, 28 Jul 2003 08:01:54 +0000");
}

sub test_to_email_mime_with_no_parts {
  my $p = shift;
  $p->{BODY_PLAIN} = undef;
  $p->{BODY_HTML} = undef;
  ok(defined $p->to_email_mime);
}

sub test_to_email_mime_with_plain_part {
  my $p = shift;
  $p->{BODY_PLAIN} = "plain";
  $p->{BODY_HTML} = undef;
  my $m = $p->to_email_mime;
  ok(defined $m);
  ok(($m->parts) == 1);
  is($m->body, "plain");
  is($m->content_type, "text/plain; charset=\"ISO-8859-1\"");
}

sub test_to_email_mime_with_html_part {
  my $p = shift;
  $p->{BODY_PLAIN} = undef;
  $p->{BODY_HTML} = "html";
  my $m = $p->to_email_mime;
  ok(defined $m);
  ok(($m->parts) == 1);
  is($m->body, "html");
  is($m->content_type, "text/html");
}

sub test_to_email_mime_with_two_parts {
  my $p = shift;
  $p->{BODY_PLAIN} = "plain";
  $p->{BODY_HTML} = "html";
  my $m = $p->to_email_mime;
  ok(defined $m);
  ok(($m->parts) == 2);
  is(($m->parts)[0]->body, "plain\r\n");
  is(($m->parts)[1]->body, "html\r\n");
}

use strict;
use warnings;
use Test::More;
use Email::Outlook::Message;
use Email::MIME::ContentType;
use Email::MIME::Modifier;
use IO::All;

my $dir = "./t/files/";
my @msgfiles = map { $_->name} grep /.*\.msg$/, io($dir)->all;
plan tests => 2 * scalar @msgfiles;
foreach my $msg (@msgfiles) {
  my $mime = $msg;
  $mime =~ s/\.msg$/\.eml/;
  my $target = Email::MIME->new(io($mime)->all);

  my $mail = Email::Outlook::Message->new($msg)->to_email_mime;

  is_deeply(get_parts($mail), get_parts($target),
    "Checking if body structure for $msg is the same");
  is_deeply(get_headers($mail), get_headers($target),
    "Checking if headers for $msg are the same");
}


sub get_parts {
  my $m = shift;

  my $arr = [];

  $m->walk_parts(sub {
      my $part = shift;
      my $ct = sanitize_content_type($part->content_type);

      if ($part->subparts > 0) {
	# TODO: Remove once we get the 'This is a multi ...' message in
	# there.
	push(@$arr, [$ct, get_headers($part), "ignore", $part->filename]);
      } elsif ($ct =~/message\/rfc822/) {
	my $mess = Email::MIME->new($part->body);
	my $struct = get_parts($mess);
	push(@$arr, [$ct, get_headers($part), $struct, $part->filename]);
      } else {
        my $body;
        if ($ct =~ /text/) {
          $body = $part->body_str;
        } else {
          $body = $part->body;
        }
	push(@$arr, [$ct, get_headers($part), $body, $part->filename]);
      }
    });

  return $arr;
}

sub get_headers {
  my $m = shift;

  my @names = sort grep(!/^content-id$/, map {lc $_} $m->header_names);
  my @arr = map {
    my @h = map { $_ =~ s/\s\s*/ /sg; $_ } sort $m->header($_);
    @h = map { sanitize_content_type($_) } @h if lc $_ eq 'content-type';
    @h = map { sanitize_content_disposition($_) } @h if lc $_ eq 'content-disposition';
    $_ . ": " . join "\n", @h;
  } @names;
  return \@arr;
}

sub sanitize_content_type {
  my $s = shift;
  my $ct = parse_content_type($s);
  my $at = $ct->{attributes};
  delete $at->{boundary};
  $at->{charset} = "us-ascii" unless exists $at->{charset};
  return join("; ", "$ct->{discrete}/$ct->{composite}",
    map("$_=\"$at->{$_}\"", sort keys %$at));
}

sub sanitize_content_disposition {
  my $s = shift;
  my $cd = parse_content_disposition($s);
  my $at = $cd->{attributes};
  return join("; ", map("$_=\"$at->{$_}\"", sort keys %$at));
}

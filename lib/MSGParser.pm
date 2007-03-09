package MSGParser;
use strict;
use Email::Simple;
use Email::Abstract;
use Email::MIME::Creator;
use Email::MIME::ContentType;
use Date::Format;
use OLE::Storage_Lite;
use POSIX qw(mktime);

use constant DIR_TYPE => 1;
use constant FILE_TYPE => 2;

use vars qw($skipproperties $skipheaders);
#
# Descriptions partially based on mapitags.h
#
$skipproperties = {
  # Envelope properties
  '000B' => "Conversation key?",
  '001A' => "Type of message",
  '003B' => "Sender address variant",
  '003D' => "Contains 'Re: '",
  '003F' => "'recieved by' id",
  '0040' => "'recieved by' name",
  '0041' => "Sender variant address id",
  '0042' => "Sender variant name",
  '0043' => "'recieved representing' id",
  '0044' => "'recieved representing' name",
  '0046' => "Read receipt address id",
  '0051' => "'recieved by' search key",
  '0052' => "'recieved representing' search key",
  '0053' => "Read receipt search key",
  '0064' => "Sender variant address type",
  '0065' => "Sender variant address",
  '0070' => "Conversation topic",
  '0071' => "Conversation index",
  '0075' => "'recieved by' address type",
  '0076' => "'recieved by' email address",
  '0077' => "'recieved representing' address type",
  '0078' => "'recieved representing' email address",
  '007F' => "something like a message id",
  # Recipient properties
  '0C19' => "Reply address variant",
  '0C1D' => "Reply address variant",
  '0C1E' => "Reply address type",
  # Non-transmittable properties
  '0E02' => "?Should BCC be displayed",
  '0E0A' => "sent mail id",
  '0E1D' => "Subject w/o Re",
  '0E27' => "64 bytes: Unknown",
  '0FF6' => "Index",
  '0FF9' => "Index",
  '0FFF' => "Address variant",
  # Content properties
  '1008' => "Summary or something",
  '1009' => "RTF Compressed",
  # 'Common property'
  '3001' => "Display name",
  '3002' => "Address Type",
  '300B' => "'Search key'",
  # Attachment properties
  '3702' => "Attachment encoding",
  '3703' => "Attachment extension",
  '3709' => "'Attachment rendering'", # Maybe an icon or something?
  '3713' => "Icon URL?",
  # 'Mail user'
  '3A20' => "Address variant",
  # 3900 -- 39FF: 'Address book'
  '39FF' => "7 bit display name",
  # 'Display table properties'
  '3FF8' => "Routing data?",
  '3FF9' => "Routing data?",
  '3FFA' => "Routing data?",
  '3FFB' => "Routing data?",
  # 'Transport-defined envelope property'
  '4029' => "Sender variant address type",
  '402A' => "Sender variant address",
  '402B' => "Sender variant name",
  '5FF6' => "Recipient name",
  '5FF7' => "Recipient address variant",
  # 'Provider-defined internal non-transmittable property'
  '6740' => "Unknown, binary data",
  # User defined id's
  '8000' => "Content Class",
  '8002' => "Unknown, binary data",
};

$skipheaders = {
  map { uc($_) => 1 } 
  "MIME-Version",
  "Content-Type",
  "Content-Transfer-Encoding",
  "X-Mailer",
  "X-Msgconvert",
  "X-MS-Tnef-Correlator",
  "X-MS-Has-Attach"
};

use constant ENCODING_UNICODE => '001F';
use constant KNOWN_ENCODINGS => {
    '000D' => 'Directory',
    '001F' => 'Unicode',
    '001E' => 'Ascii?',
    '0102' => 'Binary',
};

use constant MAP_ATTACHMENT_FILE => {
  '3701' => ["DATA",	    0], # Data
  '3704' => ["SHORTNAME",   1], # Short file name
  '3707' => ["LONGNAME",    1], # Long file name
  '370E' => ["MIMETYPE",    1], # mime type
  '3716' => ["DISPOSITION", 1], # disposition
};

use constant MAP_SUBITEM_FILE => {
  '1000' => ["BODY_PLAIN",	1], # Body
  '1013' => ["BODY_HTML",	1], # HTML Version of body
  '0037' => ["SUBJECT",		1], # Subject
  '0047' => ["SUBMISSION_ID",	1], # Seems to contain the date
  '007D' => ["HEAD",		1], # Full headers
  '0C1A' => ["FROM",		1], # Reply-To: Name
  '0C1E' => ["FROM_ADDR_TYPE",	1], # From: Address type
  '0C1F' => ["FROM_ADDR",	1], # Reply-To: Address
  '0E04' => ["TO",		1], # To: Names
  '0E03' => ["CC",		1], # Cc: Names
  '1035' => ["MESSAGEID",	1], # Message-Id
  '1042' => ["INREPLYTO",	1], # In reply to Message-Id
};

use constant MAP_ADDRESSITEM_FILE => {
  '3001' => ["NAME",		1], # Real name
  '3002' => ["TYPE",		1], # Address type
  '403D' => ["TYPE",		1], # Address type
  '3003' => ["ADDRESS",		1], # Address
  '403E' => ["ADDRESS",		1], # Address
  '39FE' => ["SMTPADDRESS",	1], # SMTP Address variant
};

#
# Main body of module
#

sub new {
  my $that = shift;
  my $file = shift or die "File name is required parameter";
  my $verbose = shift;

  my $self = $that->_empty_new;

  my $msg = OLE::Storage_Lite->new($file);
  my $pps = $msg->getPpsTree(1);
  $pps or die "Parsing $file as OLE file failed.";
  $self->set_verbosity(1) if $verbose;
  $self->_parse($pps);

  return $self;
}

sub _empty_new {
  my $that = shift;
  my $class = ref $that || $that;

  my $self = {
    ATTACHMENTS => [],
    ADDRESSES => [],
    VERBOSE => 0,
    HAS_UNICODE => 0,
    FROM_ADDR_TYPE => "",
  };
  bless $self, $class;
  return $self;
}

#
# Main sub: parse the PPS tree, and return 
#
sub _parse {
  my $self = shift;
  my $pps = shift or die "Internal error: No PPS tree";
  $self->_RootDir($pps);
}

sub _mime_object {
  my $self = shift;

  my ($plain, $html);
  my $bodymime;
  my $mime;

  unless ($self->{BODY_HTML} or $self->{BODY_PLAIN}) {
    $self->{BODY_PLAIN} = "";
  }
  if ($self->{BODY_PLAIN}) {
    $plain = Email::MIME->create(
      attributes => {
	content_type => "text/plain",
	charset => "ISO-8859-1",
	disposition => "inline",
	encoding => "8bit",
      },
      body => $self->{BODY_PLAIN}
    );
  }
  if ($self->{BODY_HTML}) {
    $html = Email::MIME->create(
      attributes => {
	content_type => "text/html"
      },
      body => $self->{BODY_HTML}
    );
  }

  if ($html and $plain) {
    $bodymime = Email::MIME->create(
      attributes => {
	content_type => "multipart/alternative",
	encoding => "8bit",
      },
      parts => [$plain, $html]
    );
  } elsif ($html) {
    $bodymime = $html;
  } else {
    $bodymime = $plain;
  }

  if (@{$self->{ATTACHMENTS}}>0) {
    my $mult = Email::MIME->create(
      attributes => {
	content_type => "multipart/mixed",
	encoding => "8bit",
      },
      parts => [$bodymime],
    );
    foreach my $att (@{$self->{ATTACHMENTS}}) {
      $self->_SaveAttachment($mult, $att);
    }
    $mime = $mult;
  } else {
    $mime = $bodymime;
  }

  $mime->header_set('Date', undef);
  $self->_copy_header_data($mime);

  $self->_SetHeaderFields($mime);

  return $mime;
}

# Actually output the message in mbox format
sub as_mbox {
  my $self = shift;

  # Construct From line from whatever we know.
  my $from = (
    $self->{FROM_ADDR_TYPE} eq "SMTP" ?
    $self->{FROM_ADDR} :
    'someone@somewhere'
  );
  $from =~ s/\n//g;

  # The date used here is not really important.
  my $mbox = "From $from " . scalar localtime() . "\n";
  $mbox .= $self->as_string;
  $mbox .= "\n";
  return $mbox;
}

sub as_string {
  my $self = shift;
  return $self->_mime_object->as_string;
}

sub set_verbosity {
  my ($self, $verbosity) = @_;
  defined $verbosity or die "Internal error: no verbosity level";
  $self->{VERBOSE} = $verbosity;
}

#
# Below are functions that walk the PPS tree. The *Dir functions handle
# processing the directory nodes of the tree (mainly, iterating over the
# children), whereas the *Item functions handle processing the items in the
# directory (if such an item is itself a directory, it will in turn be
# processed by the relevant *Dir function).
#

#
# RootItem: Check Root Entry, parse sub-entries.
# The OLE file consists of a single entry called Root Entry, which has
# several children. These children are parsed in the sub SubItem.
# 
sub _RootDir {
  my ($self, $pps) = @_;

  foreach my $child (@{$pps->{Child}}) {
    $self->_SubItem($child);
  }
}

sub _SubItem {
  my ($self, $pps) = @_;
  
  if ($pps->{Type} == DIR_TYPE) {
    $self->_SubItemDir($pps);
  } elsif ($pps->{Type} == FILE_TYPE) {
    $self->_SubItemFile($pps);
  } else {
    warn "Unknown entry type: $pps->{Type}";
  }
}

sub _SubItemDir {
  my ($self, $pps) = @_;

  $self->_GetOLEDate($pps);

  my $name = $self->_GetName($pps);

  if ($name =~ /__recip_version1 0_ /) { # Address of one recipient
    $self->_AddressDir($pps);
  } elsif ($name =~ '__attach_version1 0_ ') { # Attachment
    $self->_AttachmentDir($pps);
  } else {
    $self->_UnknownDir($self->_GetName($pps));
  }
}

sub _SubItemFile {
  my ($self, $pps) = @_;

  my $name = $self->_GetName($pps);
  my ($property, $encoding) = $self->_ParseItemName($name);

  $self->_MapProperty($self, $pps->{Data}, $property, MAP_SUBITEM_FILE)
    or $self->_UnknownFile($name);
}

sub _AddressDir {
  my ($self, $pps) = @_;

  my $address = {
    NAME	=> undef,
    ADDRESS	=> undef,
    TYPE	=> "",
  };
  foreach my $child (@{$pps->{Child}}) {
    $self->_AddressItem($child, $address);
  }
  push @{$self->{ADDRESSES}}, $address;
}

sub _AddressItem {
  my ($self, $pps, $addr_info) = @_;

  my $name = $self->_GetName($pps);

  # DIR Entries: There should be none.
  if ($pps->{Type} == DIR_TYPE) {
    $self->_UnknownDir($name);
  } elsif ($pps->{Type} == FILE_TYPE) {
    my ($property, $encoding) = $self->_ParseItemName($name);
    $self->_MapProperty($addr_info, $pps->{Data}, $property,
      MAP_ADDRESSITEM_FILE) or $self->_UnknownFile($name);
  } else {
    warn "Unknown entry type: $pps->{Type}";
  }
}

sub _AttachmentDir {
  my ($self, $pps) = @_;

  my $attachment = {
    SHORTNAME	=> undef,
    LONGNAME	=> undef,
    MIMETYPE	=> 'application/octet-stream',
    ENCODING	=> 'base64',
    DISPOSITION	=> 'attachment',
    DATA	=> undef
  };
  foreach my $child (@{$pps->{Child}}) {
    $self->_AttachmentItem($child, $attachment);
  }
  if ($attachment->{MIMETYPE} eq 'multipart/signed') {
    $attachment->{ENCODING} = '8bit';
  }
  push @{$self->{ATTACHMENTS}}, $attachment;
}

sub _AttachmentItem {
  my ($self, $pps, $att_info) = @_;

  my $name = $self->_GetName($pps);

  my ($property, $encoding) = $self->_ParseItemName($name);

  if ($pps->{Type} == DIR_TYPE) {

    if ($property eq '3701') {	# Nested MSG file
      my $msgp = $self->_empty_new();
      $msgp->_parse($pps);
      my $data = $msgp->as_string;
      $att_info->{DATA} = $data;
      $att_info->{MIMETYPE} = 'message/rfc822';
      $att_info->{ENCODING} = '8bit';
    } else {
      $self->_UnknownDir($name);
    }

  } elsif ($pps->{Type} == FILE_TYPE) {
    $self->_MapProperty($att_info, $pps->{Data}, $property,
      MAP_ATTACHMENT_FILE) or $self->_UnknownFile($name);
  } else {
    warn "Unknown entry type: $pps->{Type}";
  }
}

sub _MapProperty {
  my ($self, $hash, $data, $property, $map) = @_;

  defined $property or return 0;
  my $arr = $map->{$property} or return 0;

  if ($arr->[1]) {
    $data =~ s/\000//g;
    $data =~ s/\r\n/\n/sg;
  }
  $hash->{$arr->[0]} = $data;

  return 1;
}

sub _UnknownDir {
  my ($self, $name) = @_;

  if ($name eq '__nameid_version1 0') {
    $self->{VERBOSE}
      and warn "Skipping DIR entry $name (Introductory stuff)\n";
    return;
  }
  warn "Unknown DIR entry $name\n";
}

sub _UnknownFile {
  my ($self, $name) = @_;

  if ($name eq '__properties_version1 0') {
    $self->{VERBOSE}
      and warn "Skipping FILE entry $name (Properties)\n";
    return;
  }

  my ($property, $encoding) = $self->_ParseItemName($name);
  unless (defined $property) {
    warn "Unknown FILE entry $name\n";
    return;
  }
  if ($skipproperties->{$property}) {
    $self->{VERBOSE}
      and warn "Skipping property $property ($skipproperties->{$property})\n";
    return;
  } elsif ($property =~ /^80/) {
    $self->{VERBOSE}
      and warn "Skipping property $property (user-defined property)\n";
    return;
  } else {
    warn "Unknown property $property\n";
    return;
  }
}

#
# Helper functions
#

sub _GetName {
  my ($self, $pps) = @_;
  return $self->_NormalizeWhiteSpace(OLE::Storage_Lite::Ucs2Asc($pps->{Name}));
}

sub _NormalizeWhiteSpace {
  my ($self, $name) = @_;
  $name =~ s/\W/ /g;
  return $name;
}

sub _GetOLEDate {
  my ($self, $pps) = @_;
  unless (defined ($self->{OLEDATE})) {
    # Make Date
    my $datearr;
    $datearr = $pps->{Time2nd};
    $datearr = $pps->{Time1st} unless($datearr);
    $self->{OLEDATE} = $self->_FormatDate($datearr) if $datearr;
  }
}

sub _FormatDate {
  my ($self, $datearr) = @_;

  # TODO: This is a little convoluted. Directly using strftime didn't seem
  # to work.
  my $datetime = mktime(@$datearr);
  return time2str("%a, %d %h %Y %X %z", $datetime);
}

# If we didn't get the date from the original header data, we may be able
# to get it from the SUBMISSION_ID:
# It seems to have the format of a semicolon-separated list of key=value
# pairs. The key l has a value with the format:
# <SERVER>-<DATETIME>Z-<NUMBER>, where DATETIME is the date and time in
# the format YYMMDDHHMMSS.
sub _SubmissionIdDate {
  my $self = shift;

  my $submission_id = $self->{SUBMISSION_ID} or return undef;
  $submission_id =~ m/l=.*-(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)Z-.*/
    or return undef;
  my $year = $1;
  $year += 100 if $year < 20;
  return $self->_FormatDate([$6,$5,$4,$3,$2-1,$year]);
}

sub _ParseItemName {
  my ($self, $name) = @_;

  if ($name =~ /^__substg1 0_(....)(....)$/) {
    my ($property, $encoding) = ($1, $2);
    if ($encoding eq ENCODING_UNICODE and not ($self->{HAS_UNICODE})) {
      warn "This MSG file contains Unicode fields." 
	. " This is currently unsupported.\n";
      $self->{HAS_UNICODE} = 1;
    } elsif (not (KNOWN_ENCODINGS()->{$encoding})) {
      warn "Unknown encoding $encoding. Results may be strange or wrong.\n";
    }
    return ($property, $encoding);
  } else {
    return (undef, undef);
  }
}

sub _SaveAttachment {
  my ($self, $mime, $att) = @_;

  my $mt = parse_content_type($att->{MIMETYPE});
  my $m = Email::MIME->create(
    attributes => {
      content_type => "$mt->{discrete}/$mt->{composite}",
      %{$mt->{attributes}},
      encoding => $att->{ENCODING},
      name => ($att->{LONGNAME} ? $att->{LONGNAME} : $att->{SHORTNAME}),
      disposition => $att->{DISPOSITION},
    },
    body => $att->{DATA});
  $mime->parts_add([$m]);
}

sub _SetAddressPart {
  my ($self, $adrname, $partname, $data) = @_;

  my $address = $self->{ADDRESSES}->{$adrname};
  $data =~ s/\000//g;
  #warn "Processing address data part $partname : $data\n";
  if (defined ($address->{$partname})) {
    if ($address->{$partname} eq $data) {
      warn "Skipping duplicate but identical address information for"
      . " $partname\n" if $self->{VERBOSE};
    } else {
      warn "Address information $partname inconsistent:\n";
      warn "    Original data: $address->{$partname}\n";
      warn "    New data: $data\n";
    }
  } else {
    $address->{$partname} = $data;
  }
}

# Set header fields
sub _AddHeaderField {
  my ($self, $mime, $fieldname, $value) = @_;

  my $oldvalue = $mime->header($fieldname);
  return if $oldvalue;
  $mime->header_set($fieldname, $value) if $value;
}

sub _Address {
  my ($self, $tag) = @_;
  my $name = $self->{$tag} || "";
  my $address = $self->{$tag . "_ADDR"} || "";
  return "$name <$address>";
}

# Find SMTP addresses for the given list of names
sub _ExpandAddressList {
  my ($self, $names) = @_;

  my $addresspool = $self->{ADDRESSES};
  my @namelist = split /; */, $names;
  my @result;
  name: foreach my $name (@namelist) {
    foreach my $address (@$addresspool) {
      if ($name eq $address->{NAME}) {
	my $addresstext = $address->{NAME} . " <";
	if (defined ($address->{SMTPADDRESS})) {
	  $addresstext .= $address->{SMTPADDRESS};
	} elsif ($address->{TYPE} eq "SMTP") {
	  $addresstext .= $address->{ADDRESS};
	}
	$addresstext .= ">";
	push @result, $addresstext;
	next name;
      }
    }
    push @result, $name;
  }
  return join ", ", @result;
}

# Find out if we need to construct a multipart message
sub _IsMultiPart {
  my $self = shift;

  return (
    ($self->{BODY_HTML} and $self->{BODY_PLAIN})
      or @{$self->{ATTACHMENTS}}>0
  );
}

# Copy original header data.
# Note: This should contain the Date: header.
sub _copy_header_data {
  my ($self, $mime) = @_;

  defined $self->{HEAD} or return;
  my $parsed = new Email::Simple($self->{HEAD});

  foreach my $tag (grep { !$skipheaders->{uc $_}} $parsed->header_names) {
    $mime->header_set($tag, $parsed->header($tag));
  }
}

# Set header fields
sub _SetHeaderFields {
  my ($self, $mime) = @_;

  # If we didn't get the date from the original header data, we may be able
  # to get it from the SUBMISSION_ID:
  $self->_AddHeaderField($mime, 'Date', $self->_SubmissionIdDate());

  # Third and last chance to set the Date: header; this uses the date the
  # MSG file was saved.
  $self->_AddHeaderField($mime, 'Date', $self->{OLEDATE});
  $self->_AddHeaderField($mime, 'Subject', $self->{SUBJECT});
  $self->_AddHeaderField($mime, 'From', $self->_Address("FROM"));
  #$self->_AddHeaderField($mime, 'Reply-To', $self->_Address("REPLYTO"));
  $self->_AddHeaderField($mime, 'To', $self->_ExpandAddressList($self->{TO}));
  $self->_AddHeaderField($mime, 'Cc', $self->_ExpandAddressList($self->{CC}));
  $self->_AddHeaderField($mime, 'Message-Id', $self->{MESSAGEID});
  $self->_AddHeaderField($mime, 'In-Reply-To', $self->{INREPLYTO});
}


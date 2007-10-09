use 5.006;
use ExtUtils::MakeMaker;
WriteMakefile(
  'NAME'	    => 'Email::Outlook::Message',
  'AUTHOR'	    => 'Matijs van Zuijlen',
  'VERSION_FROM'    => 'lib/Email/Outlook/Message.pm',
  'PREREQ_PM'	=> {
    Carp => 0,
    Date::Format => 2.22,
    Email::MIME::ContentType => 1.014,
    Email::MIME::Modifier => 1.442,
    Email::MIME::Creator => 1.454,
    Email::Simple => 2.003,
    OLE::Storage_Lite => 0.14,
    POSIX => 0,
    ## These we need for testing
    IO::All => 0,
    Test::More => 0,
  },
);

1;

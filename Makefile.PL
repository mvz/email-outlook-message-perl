use 5.006;
use ExtUtils::MakeMaker;
WriteMakefile(
  'NAME'	    => 'MSGConvert',
  'AUTHOR'	    => 'Matijs van Zuijlen',
  'VERSION_FROM'    => 'msgconvert.pl',
  'EXE_FILES'	    => [ 'msgconvert.pl', 'oledump.pl' ],
  'PREREQ_PM'	=> {
    OLE::Storage_Lite => 0,
    MIME::Tools => 0,
    MIME::Entity => 0,
    MIME::Parser => 0,
    Date::Format => 0,
    POSIX => 0,
    ## These we need for testing
    #Test::Simple => 0,
    #File::Find => 0,
  },
);

1;

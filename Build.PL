use Module::Build;
my $build = Module::Build->new (
  module_name => 'Email::Outlook::Message',
  dist_abstract => 'Read Outlook .msg files',
  license => 'perl',
  create_makefile_pl => 'traditional',
  requires => {
    'perl' => 5.006,
    # Core modules
    'Carp' => '0',
    'Encode' => '0',
    'POSIX' => '0',
    'Getopt::Long' => '0',
    'Pod::Usage' => '0',
    'File::Basename' => '0',
    # Others
    'Email::MIME::ContentType' => '1.014',
    'Email::MIME' => '1.923',
    'Email::Simple' => '2.102',
    'Email::Sender' => '1.3',
    'OLE::Storage_Lite' => '0.14',
    ## These we need for testing
    'IO::All' => '0',
    'Test::More' => '0',
  },
  configure_requires => {
    'Module::Build' => 0.38
  },
  script_files => [ 'script/msgconvert' ],
);
$build->create_build_script;


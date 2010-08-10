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
    # Others
    'Email::MIME::ContentType' => '1.014',
    'Email::MIME::Modifier' => '1.442',
    'Email::MIME::Creator' => '1.454',
    'Email::Simple' => '2.003',
    'OLE::Storage_Lite' => '0.14',
    ## These we need for testing
    'IO::All' => '0',
    'Test::More' => '0',
  }
);
$build->create_build_script;


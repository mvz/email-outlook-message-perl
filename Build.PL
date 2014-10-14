use 5.006;
use Module::Build;
my $build = Module::Build->new (
  module_name => 'Email::Outlook::Message',
  dist_abstract => 'Read Outlook .msg files',
  license => 'perl',
  create_makefile_pl => 'traditional',
  requires => {
    'perl' => 5.006,
    # Core modules
    'File::Basename' => '0',
    'POSIX' => '0',
    # Others - Not versioned
    'Carp' => '0',
    'Encode' => '0',
    'Getopt::Long' => '0',
    'IO::String' => '0',
    'Pod::Usage' => '0',
    # Others - Versioned
    'Email::MIME' => '1.923',
    'Email::MIME::ContentType' => '1.014',
    'Email::Sender' => '1.3',
    'Email::Simple' => '2.102',
    'OLE::Storage_Lite' => '0.14',
    # Needed for testing
    'IO::All' => '0',
    'Test::More' => '0',
  },
  configure_requires => {
    'Module::Build' => 0.38
  },
  script_files => [ 'script/msgconvert' ],
  meta_merge => {
    resources => {
      repository => 'https://github.com/mvz/email-outlook-message-perl'
    }
  },
);
$build->create_build_script;


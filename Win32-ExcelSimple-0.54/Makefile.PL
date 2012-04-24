use lib 'inc';
use Devel::AssertOS qw[MSWin32 Cygwin];

use strict;
use warnings;



use ExtUtils::MakeMaker 6.30;

use Win32::OLE;
my $Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new( 'Excel.Application', 'Quit' );
die "install excel first" unless $Excel;
my %WriteMakefileArgs = (
  "ABSTRACT" => "a wrap of Win32::OLE excel\r",
  "AUTHOR" => "xiaoyafeng <xyf.xiao\@gmail.com>",
  "BUILD_REQUIRES" => {},
  "CONFIGURE_REQUIRES" => {
    "ExtUtils::MakeMaker" => "6.30"
  },
  "DISTNAME" => "Win32-ExcelSimple",
  "EXE_FILES" => [],
  "LICENSE" => "perl",
  "NAME" => "Win32::ExcelSimple",
  "PREREQ_PM" => {
    "Spreadsheet::Read" => 0,
    "Try::Tiny" => 0,
    "Win32::OLE" => "0.17"
  },
  "VERSION" => "0.54",
  "test" => {
    "TESTS" => "t/*.t"
  }
);


unless ( eval { ExtUtils::MakeMaker->VERSION(6.56) } ) {
  my $br = delete $WriteMakefileArgs{BUILD_REQUIRES};
  my $pp = $WriteMakefileArgs{PREREQ_PM};
  for my $mod ( keys %$br ) {
    if ( exists $pp->{$mod} ) {
      $pp->{$mod} = $br->{$mod} if $br->{$mod} > $pp->{$mod};
    }
    else {
      $pp->{$mod} = $br->{$mod};
    }
  }
}

delete $WriteMakefileArgs{CONFIGURE_REQUIRES}
  unless eval { ExtUtils::MakeMaker->VERSION(6.52) };

WriteMakefile(%WriteMakefileArgs);




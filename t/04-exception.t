use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple;
use File::Basename;
use File::Spec::Functions;

my $path = dirname(__FILE__);
   $path = Win32::GetFullPathName($path);
my $abs_file = catfile($path, 'test.xlsx');
my $es = Win32::ExcelSimple->new($abs_file);
my $sheet_h = $es->open_sheet('Report');
   is($sheet_h->write_cell(), undef, "test undef");
   
   is($sheet_h->write_row(), undef, "test write row undef");
   is($sheet_h->read(), undef,  "test read undef");
$sheet_h->write_cell(2,1,undef);
is($sheet_h->read_cell(2,1), 1, "don't overwrite cell B1 with write_cell");
$sheet_h->write_col(2,1, []);
is($sheet_h->read_cell(2,1), 1, "don't overwrite cell B1 with write_col");
$sheet_h->write_row(2,1,[]);
is($sheet_h->read_cell(2,1), 1, "don't overwrite cell B1 with write_row");
done_testing;


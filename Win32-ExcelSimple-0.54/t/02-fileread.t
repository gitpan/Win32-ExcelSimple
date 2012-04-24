use strict;
use Test::More;
use Test::Harness;
use Win32::ExcelSimple;
use File::Basename;
use File::Spec::Functions;

my $path = dirname($0);
   $path = Win32::GetFullPathName($path);
my $abs_file = catfile($path, 'test.xls');
my $es = Win32::ExcelSimple->new($abs_file);
my $sheet_h = $es->open_sheet('Report');
is($sheet_h->get_last_col(), 7, "read last col");
is($sheet_h->get_last_row(), 19, "read last row");
is($sheet_h->read_cell(3,19),  'ok', "read data from cell B1");
is_deeply($sheet_h->read(2,1,4,1),  [1,2,3], "read data from a Range");
is_deeply($sheet_h->read(2,1,2,4), [1,1,1,1], "read data from a range");
done_testing;


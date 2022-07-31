use utf8;
use strict;
use warnings;
use Test::More;
use Excel::ValueWriter::XLSX;
use Archive::Zip;

my $tmpl = "tst_remove_sheet.xlsx";
my $filename = 'tst2.xlsx';


# build an XLSX file
my $writer = Excel::ValueWriter::XLSX->new(template => $tmpl,
                                           sheets_to_remove => [qw/foo/]);

# 1st sheet, plain values and dates
$writer->add_sheet(s1 => à_table => [[qw/foo bar barbar gig/],
                                     [1, 2],
                                     [3, undef, 0, 4],
                                     [qw(01.01.2022 19.12.1999 2022-3-4 12/30/1998)],
                                     [qw(01.01.1900 28.02.1900 01.03.1900)],
                                     [qw/bar foo/]]);


# save the worksheet
$writer->save_as($filename);


# end of tests
# done_testing;




use utf8;
use strict;
use warnings;
use Test::More;
use Excel::ValueWriter::XLSX;
use Archive::Zip;

my $tmpl     = "tst_pivot.xlsx";
my $filename = 'tst3.xlsx';


# build an XLSX file
my $writer = Excel::ValueWriter::XLSX->new(template => $tmpl,
                                           sheets_to_remove => [qw/foo/]);


$writer->add_sheet(s1 => à_table => [[qw/foo bar barbar gig/],
                                     [1, 2],
                                     [3, undef, 0, 4],
                                     [qw(01.01.2022 19.12.1999 2022-3-4 12/30/1998)],
                                     [qw(01.01.1900 28.02.1900 01.03.1900)],
                                     [qw/bar foo/]]);


$writer->add_sheet(foo => tab_foo => [qw/nom val/], [[a => 1],
                                                     [b => 2],
                                                     [b => 5],
                                                     [c => 9],
                                                     ]);

# save the worksheet
$writer->save_as($filename);


# end of tests
# done_testing;




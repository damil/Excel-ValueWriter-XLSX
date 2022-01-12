use 5.028;
use strict;
use warnings;

use lib "../lib";
use Excel::ValueWriter::XLSX;

my $writer = Excel::ValueWriter::XLSX->new;

$writer->add_sheet(s1 =>      tabt1 => [[qw/foo bar barbar gig/],
                                        [1, 2],
                                        [3, undef, 0, 4],
                                        [qw(01.01.2022 19.12.1999 2022-3-4 12/30/1998)],
                                        [qw(01.01.1900 28.02.1900 01.03.1900)],
                                        [qw/bar foo/]]);
$writer->add_sheet('FEUILLE', undef,  [[qw/aa bb cc dd/], [45, 56], [qw/il était une bergère/], [99, 33, 33]]);


my $random_rows = do {my $count = 500; sub {$count-- == 500 ? [map {"h$_"} 1 .. 300] :
                                            $count          ? [map {rand()} 1 .. 300] : undef}};

$writer->add_sheet(RAND => rand => $random_rows);


$writer->save_as('foo.xlsx');

system "start foo.xlsx";





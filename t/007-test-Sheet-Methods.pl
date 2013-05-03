# -*- cperl -*-
# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

my $cntr = 1;
BEGIN { $| = 1; print "1..2\n"; }
END { print "not ok $cntr\n" unless $cntr == 0 }
use XLDB::OFile;
print "ok 1\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

# check method new
$cntr = 2;
my $ofile = XLDB::OFile->new( filename => "t/testfile.xlsx", 
			      title    => "test file", 
			      author   => "Gordon Flash", 
			      company  => "Boogy Warehouses", 
			      division => "Purchase division", 
			      toolname => "Cool Tool v1.3" );

$ofile->makeSheet( 'Income', 'L' );
$ofile->makeSheet( 'Balance', 'P' );

my $balancesheet = $ofile->sheet( 'Balance' );
$balancesheet->write( 10, 1, 'HAHA!', 'RI1x' );

my $incomesheet = $ofile->sheet( 'Income' );
$incomesheet->merge_range( 2, 2, 4, 5, 'JIJI', 'Cb' );

# without this things go wrong at destruction time
$ofile->close();
# can't do the following as the excel format sucks:
# it zips date/time info into files...
# my $out = `diff t/testfile.xlsx t/testfile.xlsx.ref`;

print "ok $cntr\n";


$cntr = 0;

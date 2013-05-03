# -*- cperl -*-
# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

my $cntr = 1;
BEGIN { $| = 1; print "1..5\n"; }
END { print "not ok $cntr\n" unless $cntr == 0 }
use XLDB::IFile;
print "ok 1\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

# check method new
$cntr = 2;
my $xldb = XLDB::IFile->new();
print "ok $cntr\n";

# check method open
$cntr = 3;
$xldb->open( 't/map1.xlsx' );
print "ok $cntr\n";

# check if sheets can be found
$cntr = 4;
for my $sheet ( qw( SomeData Blad2 ) ) {
  unless( defined( $xldb->sheet( $sheet ) ) ) {
    warn( "Error: did not find sheet '$sheet'\n" );
    print 'not ';
  }
}
print "ok $cntr\n";

# check if bogus sheets are not found
$cntr = 5;
for my $sheet ( qw( Nonexistentsheet ) ) {
  if( defined( $xldb->sheet( $sheet ) ) ) {
    warn( "Error: did find bogus sheet '$sheet'\n" );
    print 'not ';
  }
}
print "ok $cntr\n";

$cntr = 0;

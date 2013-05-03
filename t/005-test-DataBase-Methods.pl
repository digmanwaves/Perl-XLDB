# -*- cperl -*-
# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

my $cntr = 1;
BEGIN { $| = 1; print "1..15\n"; }
END { print "not ok $cntr\n" unless $cntr == 0 }
use XLDB::IFile;
use XLDB::DataBase;
use Data::Dumper;
print "ok 1\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

my $qualifier;

# check method new
$cntr = 2;
$qualifier = '';
my $xldbfile = XLDB::IFile->new();
$xldbfile->open( 't/map1.xlsx' );
my $xldb = XLDB::DataBase->new();
$xldb->connect( $xldbfile->sheet( 'SomeData' ) );
$qualifier = 'not ' unless( defined( $xldb ) );
print $qualifier . "ok $cntr\n";

# check numeric property reading
$cntr = 3;
$qualifier = '';
my %numprops = ( MinRow => 0,
		 MaxRow => 11,
		 MinCol => 0,
		 MaxCol => 9 );
while( my ( $prop, $val ) = each %numprops ) {
  my $rval = $xldb->property( $prop );
  unless( defined( $rval ) ) {
    $qualifier = 'not ';
    warn( "\nError: did not find property '$prop'\n" );
  }
  unless( $rval == $val ) {
    $qualifier = 'not ';
    warn( "Error: failed to read numeric property $prop correctly " .
	  "(expected value '$val', but found '$rval' instead.\n" );
  }
}
print $qualifier . "ok $cntr\n";


# check text property reading
$cntr = 4;
$qualifier = '';
my %txtprops = ( Name => 'SomeData' );
while( my ( $prop, $val ) = each %txtprops ) {
  my $rval = $xldb->property( $prop );
  unless ( defined( $rval ) ) {
    $qualifier = 'not ';
    warn( "Error: did not find property '$prop'\n" );
  }
  unless ( $rval eq $val ) {
    $qualifier = 'not ';
    warn( "Error: failed to read text property $prop correctly " .
	  "(expected value '$val', but found '$rval' instead.\n" );
  }
}
print $qualifier . "ok $cntr\n";

# check presence of cells
$cntr = 5;
$qualifier = '';
my $rval = $xldb->property( 'Cells' );
unless ( defined( $rval ) ) {
  $qualifier = 'not ';
  warn( "Error: did not find Cells\n" );
}
unless( ref( $rval ) eq 'ARRAY' ) {
  $qualifier = 'not ';
  warn( "Error: Cells do not seem to be present in ARRAY\n" );
}
print $qualifier . "ok $cntr\n";

# check header reading
$cntr = 6;
$qualifier = '';
$xldb->_readHeader(
		    "Sentinel",
		    { '^Sent'       => 'SENTINEL',
		      '^Number$'    => 'NR',
		      '^Name$'      => 'NAME',
		      '^Color$'     => 'COLOR',
		      '^Order$'     => 'ORDER',
		      '^XL(.*)$'    => 'XL-',
		      '^Percentage' => 'PRO',
		    }
		   );
my %headerinfo = ( 'XL-2'       => [ 7, 'XL-2' ],
		   'COLOR'      => [ 4, 'Color' ],
		   'XL-another' => [ 8, 'XL-another' ],
		   'ORDER'      => [ 5, 'Order' ],
		   'NAME'       => [ 1, 'Name' ],
		   'XL-1'       => [ 6, 'XL-1' ],
		   'SENTINEL'   => [ 0, 'Sent' ] );
while( my ( $label, $info ) = each %headerinfo ) {
  unless ( $xldb->{header}->{$label}->{COLNR} == $info->[0] ) {
    $qualifier = 'not ';
    warn( "Error: false column number in header (col with label = '$label')\n" );
  }
  unless ( $xldb->{header}->{$label}->{LABEL} eq $info->[1] ) {
    $qualifier = 'not ';
    warn( "Error: false label information in header (col with label = '$label'\n" );
  }
}
print $qualifier . "ok $cntr\n";

# check numeric property reading
$cntr = 7;
$qualifier = '';
%numprops = ( HeaderRow => 5 );
while( my ( $prop, $val ) = each %numprops ) {
  my $rval = $xldb->property( $prop );
  unless( defined( $rval ) )  {
    $qualifier = 'not ';
    warn( "Error: did not find property '$prop'\n" );
  }
  unless ( $rval == $val ) {
    $qualifier = 'not ';
    warn( "Error: failed to read numeric property $prop correctly " .
	  "(expected value '$val', but found '$rval' instead.\n" );
  }
}
print $qualifier . "ok $cntr\n";


# check line reading
$cntr = 8;
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $qualifier = 'not ' unless( defined( $line ) );
}
print $qualifier . "ok $cntr\n";

# check line reading and mandatory field checking
$cntr = 9;
$xldb->clearLogs();
my $expectedstring = "Error on line 8: missing mandatory field 'XL-1' on arbitrary line\nError on line 11: missing mandatory field 'XL-1' on arbitrary line";
my $detectedstring;
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkMandatoryFields( [ qw( COLOR NAME ) ], $line, "on arbitrary line" );
  $xldb->checkMandatoryFields( [ qw( XL-1 ) ], $line, "on arbitrary line" );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";

# check line reading and superflous field checking
$cntr = 10;
$expectedstring = "Warning on line 7: superfluous field 'XL-1' on arbitrary line (you'd better remove it)\nWarning on line 9: superfluous field 'XL-1' on arbitrary line (you'd better remove it)\nWarning on line 10: superfluous field 'XL-1' on arbitrary line (you'd better remove it)\nWarning on line 12: superfluous field 'XL-1' on arbitrary line (you'd better remove it)";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkSuperfluousFields( [ qw( XL-1 ) ], $line, "on arbitrary line" );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";
# check line reading and boolean field checking
$cntr = 11;
$expectedstring = "Error on line 8: invalid boolean 'XL-another'-value '2' (must be 0 or 1)\nError on line 9: invalid boolean 'XL-another'-value '3' (must be 0 or 1)\nError on line 10: invalid boolean 'XL-another'-value '-1' (must be 0 or 1)\nError on line 12: invalid boolean 'XL-another'-value '3' (must be 0 or 1)";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkBooleanFields( [ qw( XL-2 ) ], $line );
  $xldb->checkBooleanFields( [ qw( XL-another ) ], $line );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";


# check line reading and positive integer fields
$cntr = 12;
$expectedstring = "Error on line 7: invalid 'Number'-value '234.345' (must be a positive integer value)\nError on line 9: invalid 'Number'-value '3451.23' (must be a positive integer value)\nError on line 10: invalid 'XL-another'-value '-1' (must be a positive integer value)\nError on line 10: invalid 'Number'-value '-234.34' (must be a positive integer value)";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkPositiveIntegerFields( [ qw( XL-another ) ], $line, 12 );
  $xldb->checkPositiveIntegerFields( [ qw( NR ) ], $line, 24 );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";


# check line reading and percentage fields
$cntr = 13;
$expectedstring = "Error on line 10: invalid 'Number'-value '-234.34' (must be a positive integer value)";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkPositiveRealFields( [ qw( NR ) ], $line, 4.3 );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";

# check line reading and percentage fields
$cntr = 14;

$expectedstring = "Error on line 7: invalid 'Color'-value 'red' (must be a positive integer value)\nError on line 8: invalid 'XL-another'-value '2' (must be a value below or equal to 100%)\nError on line 8: invalid 'Color'-value 'blue' (must be a positive integer value)\nError on line 9: invalid 'Percentage'-value '1.9' (must be a value below or equal to 100%)\nError on line 9: invalid 'XL-another'-value '3' (must be a value below or equal to 100%)\nError on line 9: invalid 'Color'-value 'yellow' (must be a positive integer value)\nError on line 10: invalid 'XL-another'-value '-1' (must be a positive integer value)\nError on line 10: invalid 'Color'-value 'orange' (must be a positive integer value)\nError on line 11: invalid 'Percentage'-value '-0.25' (must be a positive integer value)\nError on line 11: invalid 'Color'-value 'peach' (must be a positive integer value)\nError on line 12: invalid 'XL-another'-value '3' (must be a value below or equal to 100%)\nError on line 12: invalid 'Color'-value 'white' (must be a positive integer value)";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkPercentageFields( [ qw( PRO ) ], $line);
  $xldb->checkPercentageFields( [ qw( XL-2 ) ], $line );
  $xldb->checkPercentageFields( [ qw( XL-another ) ], $line );
  $xldb->checkPercentageFields( [ qw( COLOR ) ], $line );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";


# check line reading and enum fields
$cntr = 15;
$expectedstring = "Error on line 8: invalid 'Color'-value 'blue' (must be one of 'red', 'yellow', 'orange')\nError on line 10: invalid 'XL-1'-value 'je' (must be one of 'bla', 'bo')\nError on line 11: invalid 'Color'-value 'peach' (must be one of 'red', 'yellow', 'orange')\nError on line 12: invalid 'Color'-value 'white' (must be one of 'red', 'yellow', 'orange')\nError on line 12: invalid 'XL-1'-value 'ji' (must be one of 'bla', 'bo')";
$detectedstring = '';
$xldb->clearLogs();
for( my $i = $xldb->property( 'HeaderRow' ) + 1; $i <= $xldb->property( 'MaxRow' ); ++$i ) {
  my $line = $xldb->parseLine( $i );
  $xldb->checkEnumFields( [ qw( COLOR ) ], [ qw( red yellow orange ) ], $line );
  $xldb->checkEnumFields( [ qw( XL-1 ) ], [ qw( bla bo ) ], $line );
}
$detectedstring .= $xldb->errors() . $xldb->warnings();
unless ( $expectedstring eq $detectedstring )
{
  warn( "Expected: $expectedstring; Got: $detectedstring" );
  print 'not ';
}
print "ok $cntr\n";


$cntr = 0;

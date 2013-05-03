require 5.006;

use ExtUtils::MakeMaker;

WriteMakefile
  (
   NAME      => 'XLDB',
   AUTHOR    => 'Walter Daems',
   ABSTRACT  => 'Use excell files as cheap database tables',
   VERSION   => '1.00',
   LICENSE   => 'perl',
   PREREQ_PM => { 'Spreadsheet::XLSX' => 0,
		  'Encode' => 0,
		},
#   EXE_FILES => [ '' ],
   test      => { TESTS => 't/???-test-*.pl' },
);



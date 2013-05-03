require 5.006;

use ExtUtils::MakeMaker;

my $target = 'MS'; # set this to MS to avoid generating man pages with 

WriteMakefile
  (
   NAME      => 'XLDB',
   AUTHOR    => 'Walter Daems / Digital Manifold Waves (walter@digmanwaves.net)',
   ABSTRACT  => 'Read and write structured data from excel files (Office Open XML)',
   VERSION   => '1.00',
   LICENSE   => 'perl',
   PREREQ_PM => { 'Spreadsheet::XLSX' => 0,
		  'Encode' => 0,
		},
   ( $target eq 'MS' ) ? ( MAN1PODS => {}, MAN3PODS => {} ) : (),
   test      => { TESTS => 't/???-test-*.pl' },
);



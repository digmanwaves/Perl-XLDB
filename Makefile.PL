require 5.006;

use ExtUtils::MakeMaker;

$target = 'MS'; # set this to MS to avoid generating man pages with 
$version = '1.00',

WriteMakefile
  (
   NAME      => 'XLDB',
   AUTHOR    => 'Walter Daems / Digital Manifold Waves (walter@digmanwaves.net)',
   ABSTRACT  => 'Read and write structured data from excel files (Office Open XML)',
   VERSION   => $version,
   LICENSE   => 'perl',
   PREREQ_PM => { 'Spreadsheet::XLSX' => 0,
		  'Encode' => 0,
		},
   ( $target eq 'MS' ) ? ( MAN1PODS => {}, MAN3PODS => {} ) : (),
   META_ADD  => 
   { provides => {
       'XLDB::File'     => { file => 'lib/XLDB/File.pm',
			     version => $version },
       'XLDB::IFile'    => { file => 'lib/XLDB/IFile.pm',
			     version => $version },
       'XLDB::OFile'    => { file => 'lib/XLDB/OFile.pm',
			     version => $version },
       'XLDB::Sheet'    => { file => 'lib/XLDB/Sheet.pm',
			     version => $version },
       'XLDB::DataBase' => { file => 'lib/XLDB/DataBase.pm',
			     version => $version },
     }
   },
   test      => { TESTS => 't/???-test-*.pl' },
  );



# -*- cperl -*-
package XLDB::IFile;

=head1 NAME

XLDB::IFile -- excel input file containing XLDB sheets

=cut

require 5.006;
use XLDB::File;
use parent 'XLDB::File';

require Spreadsheet::XLSX;

use strict;

sub new {
  my $class = shift;
  my $self = {};
  bless( $self, $class );
  return $self;
}

sub DESTROY {
  my $self = shift;
  $self->close();
}

sub open {
  
  my $self = shift;
  my ( $filename ) = @_;

  $self->{IFilename} = $filename;
  $self->{Base} = Spreadsheet::XLSX->new( $filename );

  foreach my $sheet ( @{$self->{Base}->{Worksheet}} ) {
    $self->{Sheets}->{$sheet->{Name}} = $sheet;
  }
}

sub close {
  my $self = shift;
  delete $self->{Base};
}

1;

__END__

=head1 SEE ALSO

 --

=head1 COPYRIGHT

 CONFIDENTIAL AND PROPRIETARY (C) 2013 Walter Daems / Digital Manifold Waves

=head1 AUTHOR

 Digital Manifold Waves -- F<walter@digmanwaves.net>

=cut

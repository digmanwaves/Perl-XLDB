# -*- cperl -*-
package XLDB::File;

=head1 NAME

XLDB::File -- base class for IFile and OFile

=cut

require 5.006;

use strict;

sub name {
  my $self = shift;
  return $self->{Filename};
}

sub sheet {
  my $self = shift;
  my ( $sheetName ) = @_;
  if ( exists $self->{Sheets}->{$sheetName} ) {
    return $self->{Sheets}->{$sheetName};
  }
  else {
    return undef;
  }
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

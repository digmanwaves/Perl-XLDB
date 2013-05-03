# -*- cperl -*-
package XLDB::Sheet;

=head1 NAME

XLDB::Sheet -- Excel sheet (part of a book)

=cut

require 5.006;

require Excel::Writer::XLSX;
require Excel::Writer::XLSX::Utility;

use strict;

sub new {
  my $class = shift;
  my ( $sheetname, $orientation, $file ) = @_;

  my $self = { MINROW => 1e6,
	       MAXROW => 0,
	       MINCOL => 1e6,
	       MAXCOL => 0 };

  $self->{Sheetname} = $sheetname;
  $self->{Orientation} = $orientation;
  $self->{File} = $file;
  $self->{Sheet} = $self->{File}->{Book}->add_worksheet( $sheetname )
    or return undef;

  bless( $self, $class );
  return $self;
}

sub close {
  my $self = shift;
  my ( $tool, $company, $division ) = @_;

  my ( $sec, $min, $hour, $mday, $mon, $year ) = localtime time;
  $year += 1900; ++$mon;

  $self->{Sheet}->set_paper( 9 );
  if( $self->{Orientation} =~ /^L/ ) {
    $self->{Sheet}->set_landscape();
  }
  else {
    $self->{Sheet}->set_portrait();
  }
  $self->{Sheet}->print_area( $self->{MINROW}, $self->{MINCOL},
			      $self->{MAXROW}, $self->{MAXCOL} );
  $self->{Sheet}->fit_to_pages( 1, 0 );
  $self->{Sheet}->center_horizontally();
  $self->{Sheet}->set_header( "&L&8$self->{Sheetname}" .
			      "&C&8" .
			      "&R&8&P of &N" );
  $self->{Sheet}->set_footer( "&L&8$tool" .
			      "&C&8Created: $year-$mon-$mday - $hour:$min:$sec" .

			      "&R&8$company / $division" );
  $self->{Sheet}->hide_gridlines( 1 );
}

sub write {
  my $self = shift;
  my ( $row, $col, $content, $format ) = @_;

  $self->_register_row( $row );
  $self->_register_col( $col );

  # morph format string into correct format object
  my $formatobj = undef;
  $formatobj = $self->{File}->format( $format) if defined ($format);
  $self->{Sheet}->write( $row, $col, $content, $formatobj );
  return $self;
}

sub write_formula {
  my $self = shift;
  my ( $row, $col, $content, $format ) = @_;

  $self->_register_row( $row );
  $self->_register_col( $col );

  # morph format string into correct format object
  my $formatobj = undef;
  $formatobj = $self->{File}->format( $format) if defined ($format);
  $self->{Sheet}->write_formula( $row, $col, $content, $formatobj );
  return $self;
}

sub write_url {
  my $self = shift;
  my ( $row, $col, $link, $format, $text ) = @_;

  $self->_register_row( $row );
  $self->_register_col( $col );

  # morph format string into correct format object
  my $formatobj = undef;
  $formatobj = $self->{File}->format( $format) if defined ($format);
  $self->{Sheet}->write_url( $row, $col, $link, $formatobj, $text );
  return $self;
}

sub insert_image {
  my $self = shift;
  my ( $row, $col, $image ) = @_;

  $self->_register_row( $row );
  $self->_register_col( $col );

  $self->{Sheet}->insert_image( $row, $col, $image );
  return $self;
}

sub autofilter {
  my $self = shift;
  my ( $frow, $fcol, $lrow, $lcol ) = @_;

  $self->_register_row( $frow );
  $self->_register_col( $fcol );
  $self->_register_row( $lrow );
  $self->_register_col( $lcol );

  $self->{Sheet}->autofilter( $frow, $fcol, $lrow, $lcol );
  return $self;
}

sub set_column {
  my $self = shift;
  my ( $lcol, $rcol, $width, $format ) = @_;

  $self->_register_col( $lcol );
  $self->_register_col( $rcol );

  $self->{Sheet}->set_column( $lcol, $rcol, $width,
			      $self->{File}->format( $format ) );
}

sub set_row {
  my $self = shift;
  my ( $row, $height ) = @_;

  $self->_register_row( $row );

  $self->{Sheet}->set_row( $row, $height );
}

sub merge_range {
  my $self = shift;
  my ( $inirow, $inicol, $endrow, $endcol, $content, $format ) = @_;

  $self->_register_row( $inirow );
  $self->_register_row( $endrow );
  $self->_register_col( $inicol );
  $self->_register_col( $endcol );

  $self->{Sheet}->merge_range( $inirow, $inicol, $endrow, $endcol, $content,
			       $self->{File}->format( $format ) );
}

sub _min {
  return $_[0] < $_[1] ? $_[0] : $_[1];
}

sub _max {
  return $_[0] > $_[1] ? $_[0] : $_[1];
}

sub _register_col {
  my $self = shift;
  my ( $col ) = @_;

  $self->{MAXCOL} = _max( $self->{MAXCOL}, $col );
  $self->{MINCOL} = _min( $self->{MINCOL}, $col );
}

sub _register_row {
  my $self = shift;
  my ( $row ) = @_;

  $self->{MAXROW} = _max( $self->{MAXROW}, $row );
  $self->{MINROW} = _min( $self->{MINROW}, $row );
}

1;


__END__

=head1 SEE ALSO

 --

=head1 COPYRIGHT

 CONFIDENTIAL AND PROPRIETARY (C) 2013 Digital Manifold Waves

=head1 AUTHOR

 Digital Manifold Waves -- F<walter.daems@ua.ac.be>

=cut


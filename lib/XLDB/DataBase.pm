# -*- cperl -*-
package XLDB::DataBase;

=head1 NAME

XLDB::DataBase -- XLDB database (corresponds to a single excel sheet)

=cut

require 5.006;
use strict;
require Spreadsheet::XLSX;
require ConLogger;
require Encode;

sub new {
  my $class = shift;
  my $self = {};
  bless( $self, $class );

  $self->{header} = {};
  $self->{warnings} = [];
  $self->{errors} = [];
  return $self;
}

sub connect {
  my $self = shift;
  my ( $mySheet ) = @_;
  $self->{raw} = $mySheet;
  $self->{header} = {};

  return $self;
}

sub property {
  my $self = shift;
  my ( $propname ) = @_;

  return $self->{raw}->{$propname};
}

sub db {
  my $self = shift;
  return $self->{data};
}

sub header {
  my $self = shift;
  return $self->{header};
}

sub _removeRaw {
  my $self = shift;
  delete $self->{raw};
  if ( $self->nrOfWarnings() ) {
    ConLogger::logcont( '/!\ Warnings found - see log' );
    warn( $self->warnings() . "\n" );
  }
  if ( $self->nrOfErrors() ) {
    ConLogger::logcont( '/!\ Errors found - see log' );
    die( $self->errors() . "\n" );
  }
  else {
    ConLogger::logitem( "Database fully loaded" );
  }
}

sub _readHeader {
  my $self = shift;

  my ( $sentinel, $labels ) = @_;
  my $sheet = $self->{raw};
  my $cells = \@{$sheet->{Cells}};

  # find header row
  # (= row that contains the sentinel in the leftmost column)
  my $headerrow;
  for ( my $i = $sheet->{MinRow}; $i <= $sheet->{MaxRow}; ++$i ) {
    # print STDERR "Sentinel $i:\n";
    my $cell = $cells->[$i][$sheet->{MinCol}];
    if ( defined ( $cell )
	 and $cell->{Val} =~ /^$sentinel/ ) {
      $headerrow = $i;
      last;
    }
  }
  die( "Error: DataBase does not contain a header row\n" )
    unless( defined( $headerrow ) );
  $self->{raw}->{HeaderRow} = $headerrow;
  $self->{header} = {};


  my $colsfound = 0;
 OUTER:
  # loop over all header columns
  for ( my $j = $sheet->{MinCol}; $j <= $sheet->{MaxCol}; ++$j ) {
    # print STDERR "Headercol: $j\n";
    my $cell = $cells->[$headerrow][$j];
    next unless defined( $cell );

    my $val	= Encode::decode( 'utf8', $cell->{Val} );

    # check the available labels one by one
  INNER:
    foreach my $label (sort keys ( %$labels ) ) {
      my $acronym = $labels->{$label};
      if ( __checkHeaderEntry( $val, $label, $acronym, $self->{header}, $j ) ) {
	++$colsfound;
	next OUTER;
      }
    }
  }
  if ( $colsfound < keys( %$labels ) ) {
    my @missinglabels;
    foreach my $label ( sort keys %$labels ) {
      push( @missinglabels, $label )
	unless ( exists $self->{header}->{$labels->{$label}} );
    }
    die( "Error: DataBase has a damaged header row. I could not find the following labels: " . join( ", ", @missinglabels ) . "\n" );
  }
}

sub __checkHeaderEntry {
  my ( $colheader, $label, $acronym, $header, $j ) = @_;

  if ( $colheader =~ /$label/is ) {
    if ( defined $1 ) {
      $acronym .= $1;
      $label = $acronym;
    } else {
      $label =~ s/^\^//;
      $label =~ s/\$$//;
    }
    $header->{$acronym} = { COLNR => $j,
			    LABEL => $label };
    return 1;
  }
  return 0;
}

sub parseLine {
  my $self = shift;
  my ( $rownr ) = @_;

  $self->{raw}->{CurrentRow} = $rownr + 1;

  my $line;

  while( my ( $col, $coldesc ) = each %{$self->{header}}) {
    # print STDERR Data::Dumper->Dump( [ $col, $coldesc ], [ qw( A B ) ] );
    my $colnr = $coldesc->{COLNR};
    my $cell = $self->{raw}->{Cells}[$rownr][$colnr];
    if ( defined( $cell ) ) {
      $line->{$col} = Encode::decode( 'utf8', $cell->{Val} );
      $line->{$col} =~ s/^\s+//;
      $line->{$col} =~ s/\s+$//;
    }
  }

  return $line;
}

sub dieOnCurrentLine {
  my $self = shift;
  my ( $text ) = @_;
  push @{$self->{errors}}, "Error on line $self->{raw}->{CurrentRow}: " . $text;
}

sub warnOnCurrentLine {
  my $self = shift;
  my ( $text ) = @_;
  push @{$self->{warnings}}, "Warning on line $self->{raw}->{CurrentRow}: " . $text;
}

sub nrOfWarnings {
  my $self = shift;
  return scalar @{$self->{warnings}};
}

sub nrOfErrors {
  my $self = shift;
  return scalar @{$self->{errors}};
}

sub warnings {
  my $self = shift;
  return join( "\n", @{$self->{warnings}} );
}

sub errors {
  my $self = shift;
  return join( "\n", @{$self->{errors}} );
}

sub clearLogs {
  my $self = shift;
  $self->{errors} = [];
  $self->{warnings} = [];
}

sub checkMandatoryFields {
  my $self = shift;
  my ( $fields, $line, $linerestriction ) = @_;

  my $nrOfErrors = 0;

  foreach my $field ( @$fields ) {
    die( "Internal Error: asking for non-existing mandatory field '$field'\n" )
      unless( exists $self->{header}->{$field} );

    unless ( defined( $line->{$field} )
	     and $line->{$field} !~ /^\s*$/ ) {
      $self->dieOnCurrentLine( "missing mandatory field " .
			       "'$self->{header}->{$field}->{LABEL}' $linerestriction" );
      ++$nrOfErrors;
    }
  }
  return $nrOfErrors;
}

sub checkOptionalFieldsWithDefault {
  my $self = shift;
  my ( $fields, $line, $default ) = @_;

  foreach my $field ( @$fields ) {
    unless ( defined( $line->{$field} )
	     and $line->{$field} !~ /^\s*$/ ) {
      $line->{$field} = $default;
    }
  }
}

sub checkSuperfluousFields {
  my $self = shift;
  my ( $fields, $line, $linerestriction ) = @_;

  foreach my $field ( @$fields ) {
    die( "Internal Error: asking for non-existing superfluous field '$field'\n" )
      unless( exists $self->{header}->{$field} );

    $self->warnOnCurrentLine( "superfluous field " .
			      "'$self->{header}->{$field}->{LABEL}' $linerestriction (you'd better remove it)" )
      if ( defined( $line->{$field} )
	   and $line->{$field} !~ /^\s*$/ );
  }
}

sub checkBooleanFields {
  my $self = shift;
  my ( $fields, $line ) = @_;

  my $nrOfErrors = 0;

  foreach my $field ( @$fields ) {
    my @fieldkeys = grep { m/$field/ } keys %$line;
    # print STDERR "BOOLE: " . join( '|', @fieldkeys ) . "\n";
    #print STDERR "BOOLE: " . join( '|', keys %$line ) . "\n";

    foreach my $fieldkey ( @fieldkeys ) {
      $line->{$fieldkey} = 0 unless exists $line->{$fieldkey};
      $line->{$fieldkey} =~ s/\s//g;
      $line->{$fieldkey} = 0 unless length( $line->{$fieldkey} );
      unless ( $line->{$fieldkey} =~ /^[01]$/ ) {
	$self->dieOnCurrentLine( "invalid boolean '$self->{header}->{$fieldkey}->{LABEL}'-value '$line->{$fieldkey}' " .
				 "(must be 0 or 1)" );
	++$nrOfErrors;
      }
    }
  }
  return $nrOfErrors;
}


sub checkPositiveIntegerFields {
  my $self = shift;
  my ( $fields, $line, $default ) = @_;
  $default ||= 0;

  my $nrOfErrors = 0;

  foreach my $field ( @$fields ) {
    die( "Internal Error: asking for non-existing superfluous field '$field'\n" )
      unless( exists $self->{header}->{$field} );

    $line->{$field} = $default unless exists $line->{$field};
    $line->{$field} =~ s/\s//g;
    $line->{$field} = $default unless length( $line->{$field} );
    unless ( defined( $line->{$field} )
	     and $line->{$field} ~~ /^\d*$/ ) {
      $self->dieOnCurrentLine( "invalid '$self->{header}->{$field}->{LABEL}'-value '$line->{$field}' " .
			       "(must be a positive integer value)" );
      ++$nrOfErrors;
    }
  }
  return $nrOfErrors;
}


sub checkPositiveRealFields {
  my $self = shift;
  my ( $fields, $line, $default ) = @_;
  $default ||= 0;

  my $nrOfErrors = 0;

  foreach my $field ( @$fields ) {
    die( "Internal Error: asking for non-existing superfluous field '$field'\n" )
      unless( exists $self->{header}->{$field} );

    $line->{$field} = 0 unless exists $line->{$field};
    $line->{$field} =~ s/\s//g;
    unless ( defined( $line->{$field} )
	     and $line->{$field} ~~ /^\d*\.?\d*$/ ) {
      $self->dieOnCurrentLine( "invalid '$self->{header}->{$field}->{LABEL}'-value '$line->{$field}' " .
			       "(must be a positive integer value)" );
      ++$nrOfErrors;
    }
  }
  return $nrOfErrors;
}


sub checkPercentageFields {
  my $self = shift;
  my ( $fields, $line, $default ) = @_;
  $default ||= 0;

  my $nrOfErrors = 0;

  foreach my $field ( @$fields ) {
    my $localErrors = 0;
    if ( $localErrors = $self->checkPositiveRealFields( [ $field ], $line, $default ) ) {
      $nrOfErrors += $localErrors;
    }
    else {
      unless ( $line->{$field} <= 1.0 ) {
	$self->dieOnCurrentLine( "invalid '$self->{header}->{$field}->{LABEL}'-value '$line->{$field}' " .
				 "(must be a value below or equal to 100%)" );
	++$nrOfErrors;
      }
    }
  }
  return $nrOfErrors;
}


sub checkEnumFields {
  my $self = shift;
  my ( $fields, $enum, $line, $optional ) = @_;

  foreach my $field ( @$fields ) {
    my @fieldkeys = grep { m/$field/ } keys %$line;
    # print STDERR "ENUM: " . join( '|', @fieldkeys ) . "\n";
    foreach my $fieldkey ( @fieldkeys ) {
      for my $e ( @$enum ) {
	return if( defined( $line->{$fieldkey} ) and $line->{$fieldkey} eq $e );
      }
      $line->{$fieldkey} = '' unless defined( $line->{$fieldkey} );
      $self->dieOnCurrentLine( "invalid '$self->{header}->{$fieldkey}->{LABEL}'-value '$line->{$fieldkey}' ".
			       "(must be one of '" . join( "', '", @$enum ) . "')" )
	unless $optional;
    }
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


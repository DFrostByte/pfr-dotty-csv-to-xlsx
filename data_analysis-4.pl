#!/usr/bin/perl

use strict;
use feature 'state'; # static sub variables

use constant CSV_SEPERATOR => ', ';

use constant ROWS_PER_SAMPLE => 3;
use constant COLUMNS_PER_CLUSTER => 3;
use constant COLUMNS_PER_SOLVENT => 9;

#data array column indexes
use constant { X_COLUMN => 2, #XSN contains column number
               X_ROW => 3, #YSN contains row number
               X_DATA => 10,
               X_SAMPLE_NAME => 12, #add new elements
               X_SOLVENT => 13,
               X_DILUTION => 14 };

my @COLS_LETTERS = ( 'A'..'Z' );

###############################################################################
# subs
###############################################################################
sub get_csv_filename {
  #prompt for input CSV file
  return "$ARGV[0]"
}

sub get_xls_filename {
  #prompt for XLS save file
  return "$ARGV[0].xlsx"
  #return 'test_output.xls'
}

sub get_sample_names {
  #prompt for sample names in order
  return split ( ',', $ARGV[1] );
  #return ( '20WSJ', '20WSRG', 'HJ', 'HRG', '100J', '100RG', '40J', '40RG' );
  #return ( 'Nd', 'Ag' );
}

sub get_solvent_names {
  #prompt for solvent names in order
  return split ( ',', $ARGV[2] );
  #return ( 'Water', 'CDTA', 'Na2CO3', '4 M KOH' );
}

sub process_csv_file {
  my $csv_hashref = shift;
  my $csv_filename = shift;
  my $line = '';

  use constant TRAILING_CSV_LINES => 3;

  $csv_hashref->{'filename'} = $csv_filename;

  open ( my $fh_csv, "<", "$csv_filename" );

  chomp ( $csv_hashref->{'title'} = <$fh_csv> );
  chomp ( $csv_hashref->{'info'} = <$fh_csv> );
  @{ $csv_hashref->{'columns'} } = split ( CSV_SEPERATOR, <$fh_csv> );

  while ( chomp ( $line = <$fh_csv> ) )
  {
    push ( @{ $csv_hashref->{'data'} }, [ split ( CSV_SEPERATOR, $line ) ] );
  }

  #last few lines in CSV file after records are not used
  $#{ $csv_hashref->{'data'} } -= TRAILING_CSV_LINES;

  close ( $fh_csv );
}

sub add_data_fields {
#add 'sample name', 'solvent' and 'dilution' to each data record
#there are 12 fields per record in the CSV file
  use constant { SOLVENT_DILUTION_ROW_1 => '1',
                 SOLVENT_DILUTION_ROW_2 => '1-5',
                 SOLVENT_DILUTION_ROW_3 => '1-25' };
  my @dilutions = ( SOLVENT_DILUTION_ROW_1, SOLVENT_DILUTION_ROW_2, SOLVENT_DILUTION_ROW_3 );
  my $data_ref = shift;  #two-dimensional array of the original CSV data
  my $samples_ref = shift;
  my $solvents_ref = shift;

  foreach my $record ( @$data_ref )
  {
    use integer; #floor numbers with any fractional values

    my $sample_index = ( $record->[X_ROW] - 1 ) / ROWS_PER_SAMPLE;
    my $solvent_index = ( $record->[X_COLUMN] - 1 ) / COLUMNS_PER_SOLVENT; #9 spots across per solvent
    my $dilution_index = ( $record->[X_ROW] - 1 ) % ROWS_PER_SAMPLE; #each sample has a different dilution on each row
    my $sample_cluster = ( $record->[X_COLUMN] - 1 ) / COLUMNS_PER_CLUSTER + 1; #clusters of 3*3 spots are grouped by a sample name sub heading

    $sample_cluster %= COLUMNS_PER_SOLVENT / COLUMNS_PER_CLUSTER;
    if ( ! $sample_cluster ) { $sample_cluster = COLUMNS_PER_CLUSTER; }

    $record->[X_SAMPLE_NAME] = "$samples_ref->[$sample_index]-$sample_cluster";
    $record->[X_SOLVENT] = $solvents_ref->[$solvent_index];
    $record->[X_DILUTION] = $dilutions[$dilution_index];
  }
}

###############################################################################
# spreadsheet subs
###############################################################################
sub create_sheet_raw {
  my $data_ref = shift;
  my $workbook = shift;

  my $sheet_raw = $workbook->add_worksheet( 'Raw Data' );
  my @headings_raw = ( 'Row', 'Column', 'Sample', 'Solvent', 'Dilution', 'Data' );
  my $format_heading = $workbook->add_format( border => 1, bold => 1, bg_color => 'yellow' );

  my $col = 0;
  my $row = 0;

  # print column headings
  foreach my $heading ( @headings_raw )
  {
    $sheet_raw->write( $row, $col, $heading, $format_heading );
    $col++
  }

  # print data
  $row++;

  foreach my $record ( @$data_ref )
  {
    $col = 0;

    $sheet_raw->write( $row, $col++, $record->[X_ROW] );
    $sheet_raw->write( $row, $col++, $record->[X_COLUMN] );
    $sheet_raw->write( $row, $col++, $record->[X_SAMPLE_NAME] );
    $sheet_raw->write( $row, $col++, $record->[X_SOLVENT] );
    $sheet_raw->write( $row, $col++, $record->[X_DILUTION] );
    $sheet_raw->write( $row, $col, $record->[X_DATA] );

    $row++;
  }
}

sub update_sheet_adjust_data {
# called by create_sheet_averages() to produce a worksheet of all the averages.
# caller passes cell references (ie ='Averages'!F6').
use List::MoreUtils qw(firstidx);

  my $workbook = shift;
  my $cell_sample = shift;
  my $cell_solvent = shift;
  my $cell_average = shift;

  my $format_heading = $workbook->add_format( border => 1, bold => 1, bg_color => 'yellow' );
  my @headings = ( 'Sample', 'Solvent', 'Average', 'Row Max', 'Cutoff (%)', 'Cutoff', '% of Max' );

  state $sheet_adjust_data = undef;
  state $row = 0;
  my $col_idx = 0;
  my $col_letter = '';
  state $cell_max = '';
  state $cell_cutoff = '';

  if ( !defined ( $sheet_adjust_data ) )
  # create worksheet and setup headings and variables (max, cutoff etc.)
  {
    $sheet_adjust_data = $workbook->add_worksheet( 'Adjust Data' );

    foreach my $heading ( @headings )
    {
      $sheet_adjust_data->write( $row, $col_idx++, $heading, $format_heading );
    }
    $row++;
    $col_idx = 0;

    $col_idx = firstidx { $_ eq 'Row Max' } @headings;
    $col_letter = $COLS_LETTERS[firstidx { $_ eq 'Average' } @headings];
    $sheet_adjust_data->write( $row, $col_idx, "=max($col_letter:$col_letter)" );
    $cell_max = $COLS_LETTERS[$col_idx] . ($row + 1);

    $col_idx = firstidx { $_ eq 'Cutoff (%)' } @headings;
    $col_letter = $COLS_LETTERS[$col_idx];
    $sheet_adjust_data->write( $row, $col_idx, '5%' );
    $cell_cutoff = $col_letter . ($row + 1);

    $col_idx = firstidx { $_ eq 'Cutoff' } @headings;
    $col_letter = $COLS_LETTERS[$col_idx];
    $sheet_adjust_data->write( $row, $col_idx, "=($cell_max*$cell_cutoff)" );
    $cell_cutoff = $col_letter . ($row + 1);

    $col_idx = 0;
  }

  $sheet_adjust_data->write( $row, $col_idx++, $cell_sample );
  $sheet_adjust_data->write( $row, $col_idx++, $cell_solvent );
  $sheet_adjust_data->write( $row, $col_idx, $cell_average );

  $col_idx = firstidx { $_ eq '% of Max' } @headings;
  $col_letter = $COLS_LETTERS[firstidx { $_ eq 'Average' } @headings];
  $cell_average = $col_letter . ($row+1);
  $sheet_adjust_data->write( $row, $col_idx,
    "=if($cell_average<=$cell_cutoff,0,$cell_average/$cell_max*100)" );

  $row++;
}

sub create_sheet_averages {
  my $data_ref = shift;
  my $workbook = shift;

  my $sheet_averages = $workbook->add_worksheet( 'Averages' );
  my @headings_averages = ( 'Row', 'Column', 'Sample', 'Solvent', 'Dilution', 'Data' );
  my $format_heading = $workbook->add_format( border => 1, bold => 1, bg_color => 'yellow' );
  my $format_summary = $workbook->add_format( bold => 1, bg_color => 'yellow' );

  my @sorted_data = sort {
                           $a->[X_SOLVENT] cmp $b->[X_SOLVENT] ||
                           $a->[X_SAMPLE_NAME] cmp $b->[X_SAMPLE_NAME]
                         } @$data_ref;

  my $col = 0;
  my $row = 0;

  # print column headings
  foreach my $heading ( @headings_averages )
  {
    $sheet_averages->write( $row, $col, $heading, $format_heading );
    $col++
  }

  # for each solvent print sample names with average after each sample cluster
  my $prev_record = $sorted_data[0];
  my $first_sample = 2;
  my $last_sample = 1;
  foreach my $record ( @sorted_data )
  {
    $row++;
    $col = 0;
    $last_sample++;

    if ( $prev_record->[X_SAMPLE_NAME] cmp $record->[X_SAMPLE_NAME] )
    # if different sample, print summary average row
    {
      my ( $cell_sample, $cell_solvent, $cell_average );

      $col = 2; # sample name
      $last_sample--;
      # print average
      $sheet_averages->write( $row, $col, $prev_record->[X_SAMPLE_NAME], $format_summary );
      $cell_sample = "='Averages'!" . $COLS_LETTERS[$col] . ($row + 1);

      $sheet_averages->write( $row, ++$col, $prev_record->[X_SOLVENT], $format_summary );
      $cell_solvent = "='Averages'!" . $COLS_LETTERS[$col] . ($row + 1);

      $col +=2; # data
      $sheet_averages->write( $row, $col, "=average(f$first_sample:f$last_sample)", $format_summary );
      $cell_average = "='Averages'!" . $COLS_LETTERS[$col] . ($row + 1);

      $first_sample = $row + 2; # $row is zero-based; $first_sample is one-based
      $last_sample += 2;
      $prev_record = $record;

      $col = 0;
      $row++;

      update_sheet_adjust_data ( $workbook, $cell_sample, $cell_solvent, $cell_average );
    }

    $sheet_averages->write( $row, $col++, $record->[X_ROW] );
    $sheet_averages->write( $row, $col++, $record->[X_COLUMN] );
    $sheet_averages->write( $row, $col++, $record->[X_SAMPLE_NAME] );
    $sheet_averages->write( $row, $col++, $record->[X_SOLVENT] );
    $sheet_averages->write( $row, $col++, $record->[X_DILUTION] );
    $sheet_averages->write( $row, $col, $record->[X_DATA] );
  }
}

##############################################################################
# PROGRAM START
##############################################################################

my $filename_csv = get_csv_filename ();
my $filename_xls = get_xls_filename ();
my @solvents = get_solvent_names ();
my @samples = get_sample_names ();

my %csv = ( filename => "",
            title => "",
            info => (),
            columns => (),
            data => () );

process_csv_file ( \%csv, $filename_csv );
add_data_fields ( $csv{'data'}, \@samples, \@solvents );


###################################
#create spreadsheet
use Excel::Writer::XLSX;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( "$filename_xls" );

create_sheet_raw ( $csv{'data'}, $workbook );
create_sheet_averages ( $csv{'data'}, $workbook );

#!/usr/bin/perl

use strict;

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

sub get_csv_filename {
  #prompt for input CSV file
  return "$ARGV[0]"
}

sub get_xls_filename {
  #prompt for XLS save file
  #return "$ARGV[1]"
  return 'test_output.xls'
}

sub get_sample_names {
  #prompt for sample names in order
  #return split ( ' ', $ARGV[2] );
  return ( '20WSJ', '20WSRG', 'HJ', 'HRG', '100J', '100RG', '40J', '40RG' );
}

sub get_solvent_names {
  #prompt for solvent names in order
  #return split ( ' ', $ARGV[3] );
  return ( 'Water', 'Na2CO3', '1 M KOH', '4 M KOH' );
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

sub create_sheet_averages {
  my $data_ref = shift;
  my $workbook = shift;

  my $sheet_averages = $workbook->add_worksheet( 'Averages' );
  my @headings_raw = ( 'Row', 'Column', 'Sample', 'Solvent', 'Dilution', 'Data' );
  my $format_heading = $workbook->add_format( border => 1, bold => 1, bg_color => 'yellow' );
  my $format_average = $workbook->add_format( bold => 1, bg_color => 'yellow' );

  my @sorted_data = sort {
                           $a->[X_SOLVENT] cmp $b->[X_SOLVENT] ||
                           $a->[X_SAMPLE_NAME] cmp $b->[X_SAMPLE_NAME]
                         } @$data_ref;

  my $col = 0;
  my $row = 0;

  # print column headings
  foreach my $heading ( @headings_raw )
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
      $col = 2; # sample name
      $last_sample--;
      # print average
      $sheet_averages->write( $row, $col, $prev_record->[X_SAMPLE_NAME], $format_average );
      $sheet_averages->write( $row, ++$col, $prev_record->[X_SOLVENT], $format_average );
      $col +=2; # data
      $sheet_averages->write( $row, $col, "=average(f$first_sample:f$last_sample)", $format_average );

      $first_sample = $row + 2; # $row is zero-based; $first_sample is one-based
      $last_sample += 2;
      $prev_record = $record;

      $col = 0;
      $row++;
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

#~ foreach my $record ( @{$csv{'data'}} )
#~ {
  #~ print "$record->[X_ROW] : $record->[X_COLUMN] : $record->[X_SAMPLE_NAME]";
  #~ print " : $record->[X_SOLVENT] : $record->[X_DILUTION] : $record->[X_DATA]\n";
#~ }

###################################
#create spreadsheet
use Excel::Writer::XLSX;
use File::Temp;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( "$filename_xls" );

create_sheet_raw ( $csv{'data'}, $workbook );
create_sheet_averages ( $csv{'data'}, $workbook );

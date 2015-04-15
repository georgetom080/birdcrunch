# Utility used to transpose datasheet (from checklist kind of format to traditional datasheet format)

use strict;
use IO::Handle;
use Term::ANSIColor;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility;
use Scalar::Util::Numeric qw(isint);
use List::Util qw(min max);
STDOUT->autoflush(1);

#our $input_file = "2014.xls";
our $year = "2014";
our %camp_column    = ();
our $bird_name;
our $sci_name;
our @row_to_write = ();

parseCampChecklist2();

# Parse the datasheet
sub parseCampChecklist2 {

  my @row_contents;
  my $row_number = 0;

  my $input_file = $year.".xls";
  my $output_file = "new_".$input_file;

  # Create Output File
  my $workbook_out  = Spreadsheet::WriteExcel->new($output_file);
  my $worksheet_out = $workbook_out->add_worksheet("Datasheet");
  #@row_to_write = ("Common Name", "Scientific Name", "Reported Number", "Actual Number", "Remarks on Correction", "Year", "Date", "Site", "District", "Region");
  #$worksheet_out->write_row(0, 0, \@row_to_write);

  # Open Input File
  my $parser    = Spreadsheet::ParseExcel->new();
  my $workbook  = $parser->parse($input_file);
  
  if (!defined $workbook) {
    die $parser->error(), ": $input_file?\n";
  }
  
  for my $worksheet ($workbook->worksheets() ) {
    my ($row_min, $row_max) = $worksheet->row_range();
    my ($col_min, $col_max) = $worksheet->col_range();
   
    for my $row ($row_min .. $row_max) {
  
      for my $col ($col_min .. $col_max) {
        @row_to_write = ();
        my $cell = $worksheet->get_cell($row, $col);
        if ($cell) {
  
          # Get List of Camps
          if ($row == 0) {
            unless ($cell->value =~ /Location/i) {
              $camp_column{$col} = $cell->value;
            }
  
          # Transpose Data
          } else {
            if ($col == 0) {
              if ($cell->value) {
                $bird_name = $cell->value; 
              }
            } elsif ($col == 1) {
              if ($cell->value) {
                $sci_name = $cell->value; 
              }
            } else {
              if ($cell->value) {
                #print $camp_column{$col}.",".$bird_name.",".$sci_name.",,".$cell->value."\n"; 
                @row_to_write = ($bird_name, $sci_name, "-", $cell->value, "-", $year, "-", $camp_column{$col}, "-", "-"); 
                $worksheet_out->write_row($row_number, 0, \@row_to_write);
                $row_number++;
              }
            }
          } # Transpose Data
        }
      } # Col
      #print "\n";
    } # Row
    print color("green"), "                [DONE]\n", color("reset");
  } # Worksheets
}



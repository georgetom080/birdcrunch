# WARNING! This script is under development; use at your own risk!!

# The script can be used to replace the birdnames in sheet1 according to the corrections given in sheet2

use strict;
use IO::Handle;
use Term::ANSIColor;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility;
use Scalar::Util::Numeric qw(isint);
use List::Util qw(min max);
STDOUT->autoflush(1);

# Filenames 
our $datasheet_file  = "Sacred_Groves/Sacred_Groves_datasheet.xls";
our $birdname_col    = 8;
our $correction_file = "Sacred_Groves/To_Replace.xls";
our $output_file     = "Sacred_Groves/Sacred_Groves_datasheet_edited.xls";


our $lookup_file     = "birdcrunch_lookup_hnm.xls";
our $lookup_ioc_col         = "1";
our $lookup_cmnname_col     = "4";
our $lookup_sciname_col     = "5";
our $lookup_cmnname_bli_col = "2";
our $lookup_sciname_bli_col = "3";
our $lookup_cmnname_clm_col = "6";
our $lookup_sciname_clm_col = "7";
our $lookup_cmnname_mnp_col = "8";
our $lookup_sciname_mno_col = "9";
our $lookup_endemic_col     = "10";
our $lookup_wg_endemic_col  = "11";
our $lookup_redlist_col     = "12";
our $lookup_order_col       = "13";
our $lookup_family_col      = "14";
our $lookup_guild_col       = "15";
our $lookup_biome_col       = "16";
our $lookup_range_col       = "17";

# Variables
our %correction = ();
our %birds      = ();

createLookup();
createCorrectionTable();

makeCorrection();


# Parse the lookup table
sub makeCorrection {

	my @row_contents;

	# Open IOC File
	my $parser     = Spreadsheet::ParseExcel->new();
	my $workbookr  = $parser->parse($datasheet_file);

  my $workbookw = Spreadsheet::WriteExcel->new($output_file);
	my $worksheetw = $workbookw->add_worksheet("Datasheet");

	if (!defined $workbookr) {
		die $parser->error(), ": $datasheet_file?\n";
	}

  my $row_number = 0;
  #print "\nReading Datasheet from $datasheet_file and correcting it to $output_file\n";
	for my $worksheetr ($workbookr->worksheets() ) {

		my ($row_min, $row_max) = $worksheetr->row_range();
  	my ($col_min, $col_max) = $worksheetr->col_range();

		#print "$row_min, $row_max, $col_min, $col_max\n";
  
  	for my $row ($row_min .. $row_max) {
		  @row_contents = ();
		  for my $col ($col_min .. $col_max) {
			  my $cell = $worksheetr->get_cell($row, $col);

				if ($col == $birdname_col-1) {
				  if ($cell) {
				    my $cellvalue = $cell->value;
				    if ($correction{$cellvalue}) {
				  	  push @row_contents, $cellvalue;
				  	  push @row_contents, $correction{$cellvalue};
				  	  print "$birds{$correction{$cellvalue}}->{ioc}|$cellvalue|>|$correction{$cellvalue}\n";
				  	} else {
				  	  push @row_contents, $cellvalue;
				  	  push @row_contents, $cellvalue;
				  	  print "$birds{$cellvalue}->{ioc}|$cellvalue|=|$cellvalue\n";
				  	}
				  } else {
				    push @row_contents, "-";
				    push @row_contents, "-";
				  } # Empty birdcol
			  } else {
				  if ($cell) {
					  push @row_contents, $cell->value;
					} else {
					  push @row_contents, "-";
					}
				}
			}
      $worksheetw->write_row($row_number, 0, \@row_contents);
			$row_number++;
		}
  } # Worksheets
}




# Parse the lookup table
sub createCorrectionTable {

	# Open IOC File
	my $parser    = Spreadsheet::ParseExcel->new();
	my $workbook  = $parser->parse($correction_file);

	if (!defined $workbook) {
		die $parser->error(), ": $correction_file?\n";
	}


  #print "\nCreating Correction Table\n";
	for my $worksheet ($workbook->worksheets() ) {

		my ($row_min, $row_max) = $worksheet->row_range();
  	my ($col_min, $col_max) = $worksheet->col_range();
  
  	for my $row ($row_min .. $row_max) {
	  	my $cell1 = $worksheet->get_cell($row, 0);
  		my $cell2 = $worksheet->get_cell($row, 1);
			$correction{$cell1->value} = $cell2->value;
			#print $cell1->value."-> ".$cell2->value."\n";
		}
  } # Worksheets
}



# Parse the lookup table
sub createLookup {

  my @row_contents;

  # Open IOC File
  my $parser    = Spreadsheet::ParseExcel->new();
  my $workbook  = $parser->parse($lookup_file);

  if (!defined $workbook) {
    die $parser->error(), ": $lookup_file?\n";
  }

  for my $worksheet ($workbook->worksheets() ) {

    if ($worksheet->get_name() =~ /Lookup/i) {
      #print "\nGenerating Bird Lookup Table from $lookup_file";
  
      my ($row_min, $row_max) = $worksheet->row_range();
      my ($col_min, $col_max) = $worksheet->col_range();
      
      for my $row ($row_min .. $row_max) {
        @row_contents = ();
        for my $col ($col_min .. $col_max) {
          my $cell = $worksheet->get_cell($row, $col);
          if ($cell) {
            push @row_contents, $cell->value;
          } else {
            push @row_contents, "-";
          }
        } # Col
  
        # Skip Title row
        next, if ($row_contents[$lookup_sciname_col-1] =~ /Scientific Name/i);
  
        # Add values to Lookup Hash 
        $birds{$row_contents[$lookup_cmnname_col-1]} = {ioc        => $row_contents[$lookup_ioc_col-1],
                                                        order      => $row_contents[$lookup_order_col-1],
                                                        family      => $row_contents[$lookup_family_col-1],
                                                        scinam      => $row_contents[$lookup_sciname_col-1],
                                                        cmnnam_clm => $row_contents[$lookup_cmnname_clm_col-1],
                                                        scinam_clm => $row_contents[$lookup_sciname_clm_col-1],
                                                        endemic     => $row_contents[$lookup_endemic_col-1],
                                                        wg_endemic => $row_contents[$lookup_wg_endemic_col-1],
                                                        guild       => $row_contents[$lookup_guild_col-1],
                                                        redlist     => $row_contents[$lookup_redlist_col-1],
                                                        biome       => $row_contents[$lookup_biome_col-1],
                                                        range       => $row_contents[$lookup_range_col-1]};
        #Usage: $birds{$bird}->{ioc} etc
  
      } # Row
      #print color("green"), "    [DONE]\n", color("reset");
    } # Lookup Sheet
  } # Worksheets
}


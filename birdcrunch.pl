use strict;
use IO::Handle;
use Term::ANSIColor;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility;
use Scalar::Util::Numeric qw(isint);
STDOUT->autoflush(1);

our $lookup_file 						= "ioc211_edited.xls";
our $lookup_ioc_col 				= "1";
our $lookup_order_col 			= "2";
our $lookup_family_col 			= "3";
our $lookup_cmnname_col 		= "4";
our $lookup_sciname_col 		= "5";
our $lookup_endemic_col 		= "6";
our $lookup_guild_col 			= "8";
our $lookup_redlist_col 		= "9";
our $lookup_biome_col 			= "10";
our $lookup_range_col 			= "11";

our $datasheet_file 				= "datasheet.xls";
our $datasheet_camp_col			= "1";
our $datasheet_transect_col	= "2";
our $datasheet_time_col			= "3";
our $datasheet_bird_col			= "4";
our $datasheet_num_col			= "5";
our $datasheet_db1_col			= "6";
our $datasheet_db2_col			= "7";
our $datasheet_db3_col			= "8";
our $datasheet_db4_col			= "9";
our $datasheet_habitat_col	= "10";
our $datasheet_remarks_col	= "11";

our $checklist_file 	  		= "camp_checklist.xls";
our $checklist_camp_col			= "1";
our $checklist_bird_col			= "2";
our $checklist_sciname_col	= "3";

our $look_for_dips_file 	  = "southindia_list.xls";
our $look_for_dips_bird_col	= "1";

our $output_file						=	"analysis.xls";


our %birds = ();
our @bird_names = ();
our @bird_orders = ();
our @bird_familys = ();;
our %order_of_family = ();
our @endemics = ();
our @guilds = ();
our @redlists = ();
our @biomes = ();
our @ranges = ();
our @birds_dips = ();

our @data = ();
our @camps = ();
our @birds_in_datasheet = ();
our %camp_checklist_from_datasheet = ();
our %camp_checklist = ();

our %birdcount;
our %ordercount;
our %familycount;
our %endemiccount;
our %guildcount;
our %redlistcount;
our %biomecount;
our %rangecount;
our %camp_species_count = ();

our $abundance_sigma_row_number;
our $checklist_sigma_row_number;


# Parse the lookup table
createLookup();
print "Identified ", $#bird_names+1, " birds, ", $#bird_familys+1, " families, ", $#bird_orders+1, " orders, ", $#guilds+1, " guilds, ", $#redlists+1, " redlist classes, ",$#biomes+1, " biomes and ", $#ranges+1, " ranges.\n\n";

# Parse the datasheet
parseDatasheet();
print "Datasheet has ", $#data+1, " entries and ", $#birds_in_datasheet+1, " species (including groupings like Warbler sp) across ", $#camps+1, " camps.\n\n";

# Compare with a more comprehensive list
parseCampChecklist();
foreach my $camp (@camps) {	print "$camp: ", $#{$camp_checklist{$camp}}+1, " ";} print "species parsed from checklist.\n";

# Compare with a more comprehensive list
checkForDips();

# Sum of guilds etc
sumItAllUp();

# Create the Excelsheet
generateXLS();
print "\nDone! Please verify results in $output_file\n\n";








# Prints all the sheets in the excel file
sub generateXLS {

	my @row_to_write;
	my $row_number;
	my $col_number;
	my @row_values;
	my @header;
	my $cellpos;
	my $plnp;
	my $sumpos;
	my $from_cell;
	my $to_cell;
	my $sum_row;
	my $sumstart;
	my $sumend;
	my $formula;


  # Open new file for analysed data
	my $workbook = Spreadsheet::WriteExcel->new($output_file);
	print "Generating $output_file\n";

	# Sheet1: Transect Datasheet
	print "Sheet1:  Transect Datasheet";
	my $worksheet1 = $workbook->add_worksheet("Datasheet");

	@row_to_write = ("Camp", "Transect", "Time", "Species", "No", "<5", "5-10", "10-30", ">30", "Habitat", "Remarks");
	$worksheet1->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $datum (@data) {
		@row_to_write = ($$datum{camp}, $$datum{transect}, $$datum{time}, $$datum{bird}, $$datum{num}, $$datum{db1}, $$datum{db2}, $$datum{db3}, $$datum{db4}, $$datum{habitat}, $$datum{remarks});
		$worksheet1->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "                            [DONE]\n", color("reset");


	# Sheet2: Checklist
	print "Sheet2:  Checklist";
	my $worksheet2 = $workbook->add_worksheet("Checklist");

	@row_to_write = ("No", "IOC No", "Order", "Family", "Species", "Endemic?", "Guild", "Redlist", "Biome", "Range", @camps, "Overall");
	$worksheet2->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_in_datasheet (@bird_names) {

		if (isint $birds{$bird_in_datasheet}->{ioc}) { # Skip Warbler sp etc for checklist

			my $write = 0;
			@row_values = ();
			foreach my $camp (@camps) {
				if ($birdcount{$camp}->{$bird_in_datasheet}) {
					$camp_species_count{$camp}++;
					$write = 1;
					push @row_values, "x";
				} elsif (grep(/^$bird_in_datasheet$/, @{$camp_checklist{$camp}})) {
					$write = 1;
					push @row_values, "X";
				} else {
					push @row_values, " ";
				}
			}
			
			if (($birdcount{Total}->{$bird_in_datasheet}) or ($write)) {
				$write = 1;
				push @row_values, "x";
			} else {
				push @row_values, " ";
			}

			next, unless ($write);

			@row_to_write = ($row_number, $birds{$bird_in_datasheet}->{ioc}, $birds{$bird_in_datasheet}->{order}, $birds{$bird_in_datasheet}->{family}, $bird_in_datasheet, $birds{$bird_in_datasheet}->{endemic}, $birds{$bird_in_datasheet}->{guild}, $birds{$bird_in_datasheet}->{redlist}, $birds{$bird_in_datasheet}->{biome}, $birds{$bird_in_datasheet}->{range}, @row_values);
			$worksheet2->write_row($row_number, 0, \@row_to_write);
			$row_number++;

		} # Skip Warbler sp etc

	} #foreach bird

	@row_values = ();
	$col_number = 10;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=COUNTIF($from_cell:$to_cell,\"=x\")");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("Total", "-", "-", "-", "-", "-", "-", "-", "-", "-", @row_values);
	$worksheet2->write_row($row_number, 0, \@row_to_write);
  $checklist_sigma_row_number = $row_number;

	print color("green"), "                                     [DONE]\n", color("reset");


	# Sheet3: Abundance: Species
	print "Sheet3:  Abundance: Species";
	my $worksheet3 = $workbook->add_worksheet("Abundance_Species");

	@row_to_write = ("No", "IOC No", "Order", "Family", "Species", "Endemic?", "Guild", "Redlist", "Biome", "Range", @camps, "Overall");
	$worksheet3->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_in_datasheet (@birds_in_datasheet) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($birdcount{$camp}->{$bird_in_datasheet}) {
				push @row_values, $birdcount{$camp}->{$bird_in_datasheet};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($birdcount{Total}->{$bird_in_datasheet}) {
			push @row_values, $birdcount{Total}->{$bird_in_datasheet};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $birds{$bird_in_datasheet}->{ioc}, $birds{$bird_in_datasheet}->{order}, $birds{$bird_in_datasheet}->{family}, $bird_in_datasheet, $birds{$bird_in_datasheet}->{endemic}, $birds{$bird_in_datasheet}->{guild}, $birds{$bird_in_datasheet}->{redlist},, $birds{$bird_in_datasheet}->{biome}, $birds{$bird_in_datasheet}->{range}, @row_values);
		$worksheet3->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	# Print sigma row
	@row_values = ();
	$col_number = 10;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("Total", "-", "-", "-", "-", "-", "-", "-", "-", "-", @row_values);
	$worksheet3->write_row($row_number, 0, \@row_to_write);
  $abundance_sigma_row_number = $row_number;

	print color("green"), "                            [DONE]\n", color("reset");


	# Sheet4: Abundance: Family
	print "Sheet4:  Abundance: Family";
	my $worksheet4 = $workbook->add_worksheet("Abundance_Family");

	@row_to_write = ("No", "Family", @camps, "Overall");
	$worksheet4->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_family (@bird_familys) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($familycount{$camp}->{$bird_family}) {
				push @row_values, $familycount{$camp}->{$bird_family};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($familycount{Total}->{$bird_family}) {
			push @row_values, $familycount{Total}->{$bird_family};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $bird_family, @row_values);
		$worksheet4->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	# Print sigma row
	@row_values = ();
	$col_number = 2;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("-", "Total", @row_values);
	$worksheet4->write_row($row_number, 0, \@row_to_write);

	print color("green"), "                             [DONE]\n", color("reset");


	# Sheet5: Abundance: Order
	print "Sheet5:  Abundance: Order";
	my $worksheet5 = $workbook->add_worksheet("Abundance_Order");

	@row_to_write = ("No", "Order", @camps, "Overall");
	$worksheet5->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_order (@bird_orders) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($ordercount{$camp}->{$bird_order}) {
				push @row_values, $ordercount{$camp}->{$bird_order};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($ordercount{Total}->{$bird_order}) {
			push @row_values, $ordercount{Total}->{$bird_order};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $bird_order, @row_values);
		$worksheet5->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	# Print sigma row
	@row_values = ();
	$col_number = 2;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("-", "Total", @row_values);
	$worksheet5->write_row($row_number, 0, \@row_to_write);

	print color("green"), "                              [DONE]\n", color("reset");


	# Sheet6: Endemic Analysis
	print "Sheet6:  Birds Endemic to India";
	my $worksheet6 = $workbook->add_worksheet("India_Endemic");

	@row_to_write = ("No", "Endemic", @camps, "Overall");
	$worksheet6->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $endemic (@endemics) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($endemiccount{$camp}->{$endemic}) {
				push @row_values, $endemiccount{$camp}->{$endemic};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($endemiccount{Total}->{$endemic}) {
			push @row_values, $endemiccount{Total}->{$endemic};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $endemic, @row_values);
		$worksheet6->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "                        [DONE]\n", color("reset");


	# Sheet7: Guild Analysis
	print "Sheet7:  Guild Analysis";
	my $worksheet7 = $workbook->add_worksheet("Guild_Analysis");

	@row_to_write = ("No", "Guild", @camps, "Overall");
	$worksheet7->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $guild (@guilds) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($guildcount{$camp}->{$guild}) {
				push @row_values, $guildcount{$camp}->{$guild};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($guildcount{Total}->{$guild}) {
			push @row_values, $guildcount{Total}->{$guild};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $guild, @row_values);
		$worksheet7->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	# Print sigma row
	@row_values = ();
	$col_number = 2;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("-", "Total", @row_values);
	$worksheet7->write_row($row_number, 0, \@row_to_write);


	print color("green"), "                                [DONE]\n", color("reset");


	# Sheet8: Redlist Analysis
	print "Sheet8:  Redlisted Species";
	my $worksheet8 = $workbook->add_worksheet("Redlist");

	@row_to_write = ("No", "Redlist", @camps, "Overall");
	$worksheet8->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $redlist (@redlists) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($redlistcount{$camp}->{$redlist}) {
				push @row_values, $redlistcount{$camp}->{$redlist};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($redlistcount{Total}->{$redlist}) {
			push @row_values, $redlistcount{Total}->{$redlist};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $redlist, @row_values);
		$worksheet8->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "                             [DONE]\n", color("reset");



	# Sheet9: Biome Analysis
	print "Sheet9:  IBCN Biome Restricted Assemblage";
	my $worksheet9 = $workbook->add_worksheet("Biome_Analysis");

	@row_to_write = ("No", "Biome", @camps, "Overall");
	$worksheet9->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $biome (@biomes) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($biomecount{$camp}->{$biome}) {
				push @row_values, $biomecount{$camp}->{$biome};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($biomecount{Total}->{$biome}) {
			push @row_values, $biomecount{Total}->{$biome};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $biome, @row_values);
		$worksheet9->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "              [DONE]\n", color("reset");


	# Sheet10: Range Analysis
	print "Sheet10: IBCN Range Restricted Species";
	my $worksheet10 = $workbook->add_worksheet("Range_Analysis");

	@row_to_write = ("No", "Range", @camps, "Overall");
	$worksheet10->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $range (@ranges) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($rangecount{$camp}->{$range}) {
				push @row_values, $rangecount{$camp}->{$range};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($rangecount{Total}->{$range}) {
			push @row_values, $rangecount{Total}->{$range};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $range, @row_values);
		$worksheet10->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "                 [DONE]\n", color("reset");


	# Sheet11: Shannon Index
	print "Sheet11: Shannon and Simpson Indices";
	my $worksheet11 = $workbook->add_worksheet("Shannon_Index");

	foreach my $camp (@camps) {
		push @header, $camp;
		push @header, "n/N x log(n/N)";
		push @header, "n/N x n/N";
	}

	@row_to_write = ("No", "Species", @header);
	$worksheet11->write_row(0, 0, \@row_to_write);

  my $overall_species_count = 0;
  foreach my $bird_in_datasheet (@birds_in_datasheet) {
		$overall_species_count++, if (isint $birds{$bird_in_datasheet}->{ioc});
	}

  # Print SUM Formula (Last Row)
	$col_number = 2;
	foreach my $header (@header) {
		
	  $sumstart = xl_rowcol_to_cell(1, $col_number);
	  $sumend   = xl_rowcol_to_cell($overall_species_count, $col_number);
	  $sumpos   = xl_rowcol_to_cell($overall_species_count+1, $col_number);
		$formula  = sprintf("=SUM($sumstart:$sumend)"); 
		$worksheet11->write($sumpos, $formula);
		$col_number = $col_number + 1;
	}

	$row_number = 1;
	foreach my $bird_in_datasheet (@birds_in_datasheet) {

		if (isint $birds{$bird_in_datasheet}->{ioc}) { # Skip Warbler sp etc for checklist

			$col_number = 2;
			@row_values = ();
			foreach my $camp (@camps) {

				# Insert Birdcount
				if ($birdcount{$camp}->{$bird_in_datasheet}) {
					push @row_values, $birdcount{$camp}->{$bird_in_datasheet};
				} else {
					push @row_values, 0;
				}

				# Get locations from N and n
				$sumpos  = xl_rowcol_to_cell($overall_species_count+1, $col_number);
				$cellpos = xl_rowcol_to_cell($row_number, $col_number);

				# Insert Shannon Formula
				#$formula = "1";
				$formula = sprintf("=IF($cellpos=0, 0, ($cellpos/$sumpos)*LOG(($cellpos/$sumpos)))");
				push @row_values, $formula;

				# Insert Simpson Formula
				#$formula = "1";
				$formula = sprintf("=IF($cellpos=0, 0, ($cellpos/$sumpos)*($cellpos/$sumpos))");
				push @row_values, $formula;

				$col_number = $col_number + 3;
			}
			
			@row_to_write = ($row_number, $bird_in_datasheet, @row_values);
			#print "@row_to_write\n\n";
			$worksheet11->write_row($row_number, 0, \@row_to_write);
			$row_number++;
		}
	} # Foreach bird

	# Print Header of Summary Section
	$sum_row = $row_number+1;
	$row_number = $row_number + 3;
	@row_to_write = ("", "Camp", @camps);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

	# Print No of Species Row
	@row_values = ();
	foreach my $camp (@camps) {
		push @row_values, $camp_species_count{$camp};
	}
	@row_to_write = ("", "No of Species", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

	# Print No of Individuals roe
	@row_values = ();
	foreach my $i (0 .. $#camps) {
		$cellpos = xl_rowcol_to_cell($overall_species_count+1, 2+3*$i);
		$formula = sprintf("=$cellpos");
		push @row_values, $formula;
	}
	@row_to_write = ("", "Abundance", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

	# Fetch Shannon Index Values from the 'sum-row'
	@row_values = ();
	foreach my $i (0 .. $#camps) {
		$cellpos = xl_rowcol_to_cell($sum_row-1, 3+$i*3); 
		push @row_values, sprintf("=-1*($cellpos)");
	}
	@row_to_write = ("", "Shannon Index (log base 10)", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

	# Fetch Simpson Index Values from the 'sum-row'
	@row_values = ();
	foreach my $i (0 .. $#camps) {
		$cellpos = xl_rowcol_to_cell($sum_row-1, 4+$i*3);
		push @row_values, sprintf("=$cellpos");
	}
	@row_to_write = ("", "Simpson Index (D)", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

  # Fetch Inverse Simpson Index Values from the 'sum-row'
	@row_values = ();
	foreach my $i (0 .. $#camps) {
		$cellpos = xl_rowcol_to_cell($sum_row-1, 4+$i*3);
		push @row_values, sprintf("=1/$cellpos");
	}
	@row_to_write = ("", "Inverse Simpson Index (1/D)", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);
  $row_number++;

  # Fetch Gini-Simpson Index Values from the 'sum-row'
	@row_values = ();
	foreach my $i (0 .. $#camps) {
		$cellpos = xl_rowcol_to_cell($sum_row-1, 4+$i*3);
		push @row_values, sprintf("=1-$cellpos");
	}
	@row_to_write = ("", "Gini Simpson Index (1-D)", @row_values);
	$worksheet11->write_row($row_number, 0, \@row_to_write);

	print color("green"), "                   [DONE]\n", color("reset");


	# Sheet12: Dips
	print "Sheet12: Dips";
	my $worksheet12 = $workbook->add_worksheet("Dips");

	@row_to_write = ("No", "IOC No", "Dip");
	$worksheet12->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_dip (@birds_dips) {

		@row_to_write = ($row_number, $birds{$bird_dip}->{ioc}, $bird_dip);
		$worksheet12->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	print color("green"), "                                          [DONE]\n", color("reset");



}




# Parse the lookup table
sub createLookup {

	print "\nGenerating Bird Lookup Table from $lookup_file";
	my @row_contents;

	# Open IOC File
	my $parser    = Spreadsheet::ParseExcel->new();
	my $workbook  = $parser->parse($lookup_file);

	if (!defined $workbook) {
		die $parser->error(), ": $lookup_file?\n";
	}

	for my $worksheet ($workbook->worksheets() ) {
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
			$birds{$row_contents[$lookup_cmnname_col-1]} = {ioc 		=> $row_contents[$lookup_ioc_col-1],
																											order 	=> $row_contents[$lookup_order_col-1],
																											family 	=> $row_contents[$lookup_family_col-1],
																											scinam 	=> $row_contents[$lookup_sciname_col-1],
																											endemic	=> $row_contents[$lookup_endemic_col-1],
																											guild		=> $row_contents[$lookup_guild_col-1],
																											redlist	=> $row_contents[$lookup_redlist_col-1],
																											biome		=> $row_contents[$lookup_biome_col-1],
																											range		=> $row_contents[$lookup_range_col-1]};
			#Usage: $birds{$bird}->{ioc} etc

			# Make an array of all birds
			#print "Pushing $row_contents[$lookup_cmnname_col-1]\n";
			push @bird_names, $row_contents[$lookup_cmnname_col-1];

			# Make an array of all orders
			unless (grep(/^$row_contents[$lookup_order_col-1]$/, @bird_orders)) {
				push @bird_orders, $row_contents[$lookup_order_col-1];
			}

			# Make an array of all families
			unless (grep(/^$row_contents[$lookup_family_col-1]$/, @bird_familys)) {
				push @bird_familys, $row_contents[$lookup_family_col-1];
			}

			# Map Family to its Order
			$order_of_family{$row_contents[$lookup_family_col-1]} = $row_contents[$lookup_order_col-1];

			# Make an array of all endemics
			unless (grep(/^$row_contents[$lookup_endemic_col-1]$/, @endemics)) {
				push @endemics, $row_contents[$lookup_endemic_col-1], if (($row_contents[$lookup_endemic_col-1]) and ($row_contents[$lookup_endemic_col-1] ne "-"));
			}

			# Make an array of all guilds
			unless (grep(/^$row_contents[$lookup_guild_col-1]$/, @guilds)) {
				push @guilds, $row_contents[$lookup_guild_col-1], unless ($row_contents[$lookup_guild_col-1] eq "-" );
			}

			# Make an array of all redlists
			unless (grep(/^$row_contents[$lookup_redlist_col-1]$/, @redlists)) {
				push @redlists, $row_contents[$lookup_redlist_col-1], unless ($row_contents[$lookup_redlist_col-1] eq "-" );
			}

			# Make an array of all biomes
			unless (grep(/^$row_contents[$lookup_biome_col-1]$/, @biomes)) {
				push @biomes, $row_contents[$lookup_biome_col-1], if (($row_contents[$lookup_biome_col-1]) and ($row_contents[$lookup_biome_col-1] ne "-"));
			}

			# Make an array of all ranges
			unless (grep(/^$row_contents[$lookup_range_col-1]$/, @ranges)) {
				push @ranges, $row_contents[$lookup_range_col-1], unless ($row_contents[$lookup_range_col-1] eq "-" );
			}


		} # Row
		print color("green"), "    [DONE]\n", color("reset");
	} # Worksheets
}


# Parse the datasheet
sub parseDatasheet {

	my @all_birds_in_datasheet = ();
	my @birdname_nomatch = ();
	my @birdname_match = ();
	my @row_contents;

	print "\nReading entries from datasheet $datasheet_file";

	# Open IOC File
	my $parser    = Spreadsheet::ParseExcel->new();
	my $workbook  = $parser->parse($datasheet_file);

	if (!defined $workbook) {
		die $parser->error(), ": $datasheet_file?\n";
	}

	for my $worksheet ($workbook->worksheets() ) {
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
      next, if ($row_contents[$datasheet_transect_col-1] =~ /Transect/i);

			push @data, {	camp 			=> $row_contents[$datasheet_camp_col-1],
										transect 	=> $row_contents[$datasheet_transect_col-1],
										time 			=> $row_contents[$datasheet_time_col-1],
										bird 			=> $row_contents[$datasheet_bird_col-1],
										num 			=> $row_contents[$datasheet_num_col-1],
										db1 			=> $row_contents[$datasheet_db1_col-1],
										db2 			=> $row_contents[$datasheet_db2_col-1],
										db3 			=> $row_contents[$datasheet_db3_col-1],
										db4 			=> $row_contents[$datasheet_db4_col-1],
										habitat		=> $row_contents[$datasheet_habitat_col-1],
										remarks		=> $row_contents[$datasheet_remarks_col-1]};

			# Make an array of all birds seen atleast once in datasheet
			unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @all_birds_in_datasheet)) {
				push @all_birds_in_datasheet, $row_contents[$datasheet_bird_col-1];
			}

			# Print a warning if the entry in the datasheet has no match in the lookup
			unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @bird_names)) {
				push @birdname_nomatch, $row_contents[$datasheet_bird_col-1], unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @birdname_nomatch));
			}

			# Make an array of all camps
			unless (grep(/^$row_contents[$datasheet_camp_col-1]$/, @camps)) {
				push @camps, $row_contents[$datasheet_camp_col-1];
			}

		} # Row
		print color("green"), "           [DONE]\n", color("reset");
	} # Worksheets

	# Sorting Birds in datasheets in the IOC Order.
	foreach my $bird (@bird_names) {
		if  (grep(/^$bird$/, @all_birds_in_datasheet)) {
			push @birdname_match, $bird;
		}
	}
	push @birds_in_datasheet, @birdname_match;
	push @birds_in_datasheet, @birdname_nomatch;

	# Ideally, all birds in datasheet should have entry in the lookup table. Printring warning otherwise
	if ($#birdname_nomatch >= 0) {
		print color("yellow"), "\nWarning: The following bird names in the datasheet has no exact match in the lookup table. You may want to:\n";
		print "1. Check the spelling in the datasheet.\n";
		print "2. In cases where the entry in the datasheet is an UNID, add entries like 'Warbler sp' in the lookup.\n", color("reset");
		foreach my $birdname_nomatch (@birdname_nomatch) {
    	print "$birdname_nomatch\n";
		}
		print "\n";
	}
}



# Parse the datasheet
sub parseCampChecklist {

	my @all_birds_in_datasheet = ();
	my @birdname_nomatch = ();
	my @birdname_match = ();
	my @row_contents;

	print "\nReading entries from $checklist_file";

	# Open IOC File
	my $parser    = Spreadsheet::ParseExcel->new();
	my $workbook  = $parser->parse($checklist_file);

	if (!defined $workbook) {
		die $parser->error(), ": $checklist_file?\n";
	}

	for my $worksheet ($workbook->worksheets() ) {
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
      next, if ($row_contents[$checklist_sciname_col-1] =~ /Scientific Name/i);

			unless (grep(/^$row_contents[$checklist_bird_col-1]$/, @{$camp_checklist{$row_contents[$checklist_camp_col-1]}})) {
				#print "Camp: $row_contents[$checklist_camp_col-1]. Pushing $row_contents[$checklist_bird_col-1]\n";
				push @{$camp_checklist{$row_contents[$checklist_camp_col-1]}}, $row_contents[$checklist_bird_col-1];
			}

		} # Row
		print color("green"), "                [DONE]\n", color("reset");
	} # Worksheets

}







# Parse the datasheet
sub checkForDips {

	my @birds_look_for_dips = ();
  my @row_contents = ();

	print "\nChecking for Dips in $look_for_dips_file";

	# Open IOC File
	my $parser    = Spreadsheet::ParseExcel->new();
	my $workbook  = $parser->parse($look_for_dips_file);

	if (!defined $workbook) {
		die $parser->error(), ": $look_for_dips_file?\n";
	}

	for my $worksheet ($workbook->worksheets() ) {
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

			push @birds_look_for_dips, $row_contents[$look_for_dips_bird_col-1];

		} # Row
	} # Worksheets

	foreach my $bird (@birds_look_for_dips) {
		push @birds_dips, $bird, unless (grep(/^$bird$/, @birds_in_datasheet));
	}

	print color("green"), "               [DONE]\n\n", color("reset");
}


sub sumItAllUp {
 
	print "Analysing Data";

  # Take each line of datasheet one by one
	foreach my $datum (@data) {

		# ORDER Count (Per camp and total)
		foreach my $camp (@camps) {
			$ordercount{$camp}->{$birds{$$datum{bird}}->{order}} = $ordercount{$camp}->{$birds{$$datum{bird}}->{order}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$ordercount{Total}->{$birds{$$datum{bird}}->{order}} = $ordercount{Total}->{$birds{$$datum{bird}}->{order}} + $$datum{num};

		# FAMILY Count (Per camp and total)
		foreach my $camp (@camps) {
			$familycount{$camp}->{$birds{$$datum{bird}}->{family}} = $familycount{$camp}->{$birds{$$datum{bird}}->{family}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$familycount{Total}->{$birds{$$datum{bird}}->{family}} = $familycount{Total}->{$birds{$$datum{bird}}->{family}} + $$datum{num};

		# SPECIES Count (Per camp and total)
		foreach my $camp (@camps) {
			$birdcount{$camp}->{$$datum{bird}} = $birdcount{$camp}->{$$datum{bird}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$birdcount{Total}->{$$datum{bird}} = $birdcount{Total}->{$$datum{bird}} + $$datum{num};

		# ENDEMIC Count (Per camp and total)
		foreach my $camp (@camps) {
			$endemiccount{$camp}->{$birds{$$datum{bird}}->{endemic}} = $endemiccount{$camp}->{$birds{$$datum{bird}}->{endemic}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$endemiccount{Total}->{$birds{$$datum{bird}}->{endemic}} = $endemiccount{Total}->{$birds{$$datum{bird}}->{endemic}} + $$datum{num};

		# GUILD Count (Per camp and total)
		foreach my $camp (@camps) {
			$guildcount{$camp}->{$birds{$$datum{bird}}->{guild}} = $guildcount{$camp}->{$birds{$$datum{bird}}->{guild}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$guildcount{Total}->{$birds{$$datum{bird}}->{guild}} = $guildcount{Total}->{$birds{$$datum{bird}}->{guild}} + $$datum{num};

		# Redlist Count (Per camp and total)
		foreach my $camp (@camps) {
			$redlistcount{$camp}->{$birds{$$datum{bird}}->{redlist}} = $redlistcount{$camp}->{$birds{$$datum{bird}}->{redlist}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$redlistcount{Total}->{$birds{$$datum{bird}}->{redlist}} = $redlistcount{Total}->{$birds{$$datum{bird}}->{redlist}} + $$datum{num};

		# BIOME Count (Per camp and total)
		foreach my $camp (@camps) {
			$biomecount{$camp}->{$birds{$$datum{bird}}->{biome}} = $biomecount{$camp}->{$birds{$$datum{bird}}->{biome}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$biomecount{Total}->{$birds{$$datum{bird}}->{biome}} = $biomecount{Total}->{$birds{$$datum{bird}}->{biome}} + $$datum{num};

		# RANGE Count (Per camp and total)
		foreach my $camp (@camps) {
			$rangecount{$camp}->{$birds{$$datum{bird}}->{range}} = $rangecount{$camp}->{$birds{$$datum{bird}}->{range}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$rangecount{Total}->{$birds{$$datum{bird}}->{range}} = $rangecount{Total}->{$birds{$$datum{bird}}->{range}} + $$datum{num};

	}
	print color("green"), "                                         [DONE]\n\n", color("reset");
}

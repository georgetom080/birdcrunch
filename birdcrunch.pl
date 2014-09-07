# BirdCrunch is a Perl utility to Analyse Bird Survey Data.

# WARNING! This script is under development; use at your own risk!!

# The script can be used/extended to convert raw data from transects into various formats for analysis.
# It can also generate ebird-checklist-format csvs for uploading the data into ebird.org

# The script needs companion lookup file birdcrunch_lookup.xls, along with raw datasheet in xls format in the same folder it is excecuted from.
# Additionaly, user might also use optional camp-wise checklist file and a local list to compare and indicate dips.

# REFERENCES
# ----------
# Praveen J., Jayapal, R., Pittie, A., 2013. A Checklist of Birds of India - Non Rarities.
# IUCN Red List - BirdLife list of Threatened Birds of India (updated by IBCN till 28 July 2014)
# Western Ghat Endemics - Western Ghat Endemics-IOC 4.3
# India Endemics - according to Jaypal, R. Summary of IOC Bird List v. 2.11.
# Biome Representative and Range Specific Assemblage - According to IBA Book, IBCN.
# Feeding Guilds - Based on Raman et al. (1998) and Praveen and Nameer (2009)

# Contact: (especially if you have review comments, would like to extend this, build a gui, or make this work in Windows)
# George Tom
# georgetom080 at gmail dot com
# Kenneth Anderson Nature Society
# kans.org.in

use strict;
use IO::Handle;
use Term::ANSIColor;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility;
use Scalar::Util::Numeric qw(isint);
use List::Util qw(min max);
STDOUT->autoflush(1);

# Survey Name
our $surveyname =	"Mankulam Bird Survey, 2009.";
our $areaname   =	"Mankulam";

# Output filename
our $output_file =	"mankulam_analysis_george3.xls";

# Switch ON/OFF (1/0) the sheets to be added in the output
my $Transect_Datasheet                = 1;
my $Checklist                         = 1;
my $Abundance_Species                 = 1;
my $Abundance_Family                  = 1;
my $Abundance_Order                   = 1;
my $Birds_Endemic_to_India            = 1;
my $Birds_Endemic_to_WGhats           = 1;
my $Guild_Analysis                    = 1;
my $Redlisted_Species                 = 1;
my $IBCN_Biome_Restricted_Assemblage  = 1;
my $IBCN_Range_Restricted_Species     = 1;
my $Shannon_and_Simpson_Indices       = 1;
my $Bray_Curtis_from_Abundance        = 1;
my $Bray_Curtis_from_Checklist        = 1;
my $Dips                              = 0;
my $Ebird                             = 1;

# Filename and column positions of the datasheet from the field
our $datasheet_file         = "mankulam.xls";
our $datasheet_camp_col     = "1";
our $datasheet_date_col     = "2";
our $datasheet_transect_col = "3";
our $datasheet_time_col     = "4";
our $datasheet_bird_col     = "5";
our $datasheet_num_col      = "6";
our $datasheet_db1_col      = "8"; # Distance Band
our $datasheet_db2_col      = "9";
our $datasheet_db3_col      = "10";
our $datasheet_db4_col      = "11";
our $datasheet_habitat_col  = "7";
our $datasheet_remarks_col  = "12";
our $datasheet_observers_col= "13";


# Filename and column positions of the lookup sheet
our $lookup_file            = "birdcrunch_lookup.xls";
our $lookup_ioc_col         = "1";
our $lookup_cmnname_col     = "2";
our $lookup_sciname_col     = "3";
our $lookup_cmnname_bli_col = "4";
our $lookup_sciname_bli_col = "5";
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


# Filename and column positions of checklist file
our $checklist_file        = "camp_checklist.xls";
our $checklist_camp_col    = "1";
our $checklist_bird_col    = "2";
our $checklist_sciname_col = "3";

# Filename and column positions of local birdlist to find out dips
our $look_for_dips_file     = "southindia_list.xls";
our $look_for_dips_bird_col	= "1";

# Internal Variable Declarations
our %birds = ();
our @bird_names = ();
our @bird_orders = ();
our @bird_familys = ();;
our %order_of_family = ();
our @endemics = ();
our @wg_endemics = ();
our @guilds = ();
our @redlists = ();
our @biomes = ();
our @ranges = ();
our @birds_dips = ();
our @survey_checklist = ();
our %date = ();
our %observers = ();

our @data = ();
our @camps = ();
our %transects = ();
our @camp_transects = ();
our @camp_camp = ();
our @start_times;
our @end_times;
our @durations;
our @birds_in_datasheet = ();
our %camp_checklist_from_datasheet = ();
our %camp_checklist = ();

our %birdcount;
our %ebirdcount;
our %ordercount;
our %familycount;
our %endemiccount;
our %wg_endemiccount;
our %guildcount;
our %redlistcount;
our %biomecount;
our %rangecount;
our %camp_species_count = ();
our %ebirdofioc = ();;

our $abundance_sigma_row_number;
our $checklist_sigma_row_number;


# Parse the lookup table
createLookup();
print "Identified ", $#bird_names+1, " birds, ", $#bird_familys+1, " families, ", $#bird_orders+1, " orders, ", $#guilds+1, " guilds, ", $#redlists+1, " redlist classes, ",$#biomes+1, " biomes and ", $#ranges+1, " ranges.\n\n";

# Parse the datasheet
parseDatasheet();
print "Datasheet has ", $#data+1, " entries and ", $#birds_in_datasheet+1, " species (including groupings like Warbler sp) across ", $#camps+1, " camps.\n\n";

# Compare with a more comprehensive list
parseCampChecklist(), if ($Checklist);
foreach my $camp (@camps) {	print "$camp: ", $#{$camp_checklist{$camp}}+1, " ", if ($Checklist);} print "species parsed from checklist.\n", if ($Checklist);

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
	my $checklist;
	my $braycurtis_i;
	my $sigma_camp1;
	my $sigma_camp2;
	my $bray_curtis_coeff;

  # Open new file for analysed data
	my $workbook = Spreadsheet::WriteExcel->new($output_file);
	print "\nWriting results to $output_file\n";

if ($Transect_Datasheet) {
	# Sheet1: Transect Datasheet
	print "Sheet1:  Transect Datasheet";
	my $worksheet1 = $workbook->add_worksheet("Datasheet");

	@row_to_write = ("Camp", "Transect", "Time", "Species", "No", "<5", "5-10", "10-30", ">30", "Habitat", "Remarks", "Observers");
	$worksheet1->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $datum (@data) {
		@row_to_write = ($$datum{camp}, $$datum{transect}, $$datum{time}, $$datum{bird}, $$datum{num}, $$datum{db1}, $$datum{db2}, $$datum{db3}, $$datum{db4}, $$datum{habitat}, $$datum{remarks}, $$datum{observers});
		$worksheet1->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}
	print color("green"), "                            [DONE]\n", color("reset");
}

if ($Checklist) {
	# Sheet2: Checklist
	print "Sheet2:  Checklist";
	my $worksheet2 = $workbook->add_worksheet("Checklist");

	@row_to_write = ("No", "IOC No", "Order", "Family", "Species", "IN-Endemic", "WG-Endemic", "Guild", "Redlist", "Biome", "Range", @camps, "Overall");
	$worksheet2->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_in_datasheet (@bird_names) {

		#if (isint $birds{$bird_in_datasheet}->{ioc}) { # Skip Warbler sp etc for checklist
		if (1) {

			my $write = 0;
			@row_values = ();
			foreach my $camp (@camps) {
				if ($birdcount{$camp}->{$bird_in_datasheet}) {
					$camp_species_count{$camp}++;
					$write = 1;
					push @row_values, $birdcount{$camp}->{$bird_in_datasheet};
					#push @row_values, "x";
					push (@survey_checklist, $bird_in_datasheet), unless (grep(/^$bird_in_datasheet$/, @survey_checklist))
				} elsif (grep(/^$bird_in_datasheet$/, @{$camp_checklist{$camp}})) {
					$write = 1;
					push @row_values, "X";
					push (@survey_checklist, $bird_in_datasheet), unless (grep(/^$bird_in_datasheet$/, @survey_checklist));
				} else {
					push @row_values, " ";
				}
			}
			
			if (($birdcount{Total}->{$bird_in_datasheet}) or ($write)) {
				$write = 1;
				#push @row_values, "x";
				push @row_values, $birdcount{Total}->{$bird_in_datasheet};
			} else {
				push @row_values, " ";
			}

			next, unless ($write);

			@row_to_write = ($row_number, $birds{$bird_in_datasheet}->{ioc}, $birds{$bird_in_datasheet}->{order}, $birds{$bird_in_datasheet}->{family}, $bird_in_datasheet, $birds{$bird_in_datasheet}->{endemic}, $birds{$bird_in_datasheet}->{wg_endemic}, $birds{$bird_in_datasheet}->{guild}, $birds{$bird_in_datasheet}->{redlist}, $birds{$bird_in_datasheet}->{biome}, $birds{$bird_in_datasheet}->{range}, @row_values);
			$worksheet2->write_row($row_number, 0, \@row_to_write);
			$row_number++;

		} # Skip Warbler sp etc

	} #foreach bird

	@row_values = ();
	$col_number = 11;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		#$formula = sprintf("=COUNTIF($from_cell:$to_cell,\"=x\")");
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("Total", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", @row_values);
	$worksheet2->write_row($row_number, 0, \@row_to_write);
  $checklist_sigma_row_number = $row_number;

	print color("green"), "                                     [DONE]\n", color("reset");
}

if ($Abundance_Species) {
	# Sheet3: Abundance: Species
	print "Sheet3:  Abundance: Species";
	my $worksheet3 = $workbook->add_worksheet("Abundance_Species");

	@row_to_write = ("No", "IOC No", "Order", "Family", "Species", "IN-Endemic", "WG-Endemic", "Guild", "Redlist", "Biome", "Range", @camps, "Overall");
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

		@row_to_write = ($row_number, $birds{$bird_in_datasheet}->{ioc}, $birds{$bird_in_datasheet}->{order}, $birds{$bird_in_datasheet}->{family}, $bird_in_datasheet, $birds{$bird_in_datasheet}->{endemic}, $birds{$bird_in_datasheet}->{wg_endemic}, $birds{$bird_in_datasheet}->{guild}, $birds{$bird_in_datasheet}->{redlist},, $birds{$bird_in_datasheet}->{biome}, $birds{$bird_in_datasheet}->{range}, @row_values);
		$worksheet3->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	# Print sigma row
	@row_values = ();
	$col_number = 11;
	foreach my $i (0 .. $#camps+1) {
		$from_cell 	= xl_rowcol_to_cell(1, $col_number);
		$to_cell 		= xl_rowcol_to_cell($row_number-1, $col_number);
		$formula = sprintf("=SUM($from_cell:$to_cell)");
		push @row_values, $formula;
		$col_number++;
	}
	@row_to_write = ("Total", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", , @row_values);
	$worksheet3->write_row($row_number, 0, \@row_to_write);
  $abundance_sigma_row_number = $row_number;

	print color("green"), "                            [DONE]\n", color("reset");
}


if ($Abundance_Family) {
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
}

if ($Abundance_Order) {
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
}



if ($Birds_Endemic_to_India) {
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
}



if ($Birds_Endemic_to_WGhats) {
	# Sheet6A: Endemic Analysis
	print "Sheet6A: Birds Endemic to Western Ghats";
	my $worksheet6A = $workbook->add_worksheet("WGhats_Endemic");

	@row_to_write = ("No", "WGhats_Endemic", @camps, "Overall");
	$worksheet6A->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $wg_endemic (@wg_endemics) {

		@row_values = ();
		foreach my $camp (@camps) {
			if ($wg_endemiccount{$camp}->{$wg_endemic}) {
				push @row_values, $wg_endemiccount{$camp}->{$wg_endemic};
			} else {
				push @row_values, 0;
			}
		}
		
		if ($wg_endemiccount{Total}->{$wg_endemic}) {
			push @row_values, $wg_endemiccount{Total}->{$wg_endemic};
		} else {
			push @row_values, 0;
		}

		@row_to_write = ($row_number, $wg_endemic, @row_values);
		$worksheet6A->write_row($row_number, 0, \@row_to_write);
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
	$worksheet6A->write_row($row_number, 0, \@row_to_write);


	print color("green"), "                [DONE]\n", color("reset");
}



if ($Guild_Analysis) {
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
}

if ($Redlisted_Species) {
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
}

if ($IBCN_Biome_Restricted_Assemblage) {
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
}

if ($IBCN_Range_Restricted_Species) {
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
}

if ($Shannon_and_Simpson_Indices) {
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
}

if ($Bray_Curtis_from_Abundance) {
	# Sheet12: Bray Curtis from Abundance
	print "Sheet12: Bray Curtis from Abundance";;
	my $worksheet12 = $workbook->add_worksheet("BrayCurtis_Abundance");

	@row_to_write = ("Camp", @camps);
	$worksheet12->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $camp1 (@camps) {
		@row_values = ();
		foreach my $camp2 (@camps) {
			$braycurtis_i = 0;
			$sigma_camp1 = 0;
			$sigma_camp2 = 0;
			$bray_curtis_coeff = 0;
			#print "\n\n\nCAMPS: $camp1 $camp2\n\n";
			foreach my $bird_in_datasheet (@birds_in_datasheet) {
				
				$braycurtis_i = $braycurtis_i + 2*min($birdcount{$camp1}->{$bird_in_datasheet}, $birdcount{$camp2}->{$bird_in_datasheet});
				#print "Bird: $bird_in_datasheet\t C1: $birdcount{$camp1}->{$bird_in_datasheet}\t$birdcount{$camp2}->{$bird_in_datasheet}\tSigma: $braycurtis_i\n";
				$sigma_camp1 = $sigma_camp1 + $birdcount{$camp1}->{$bird_in_datasheet};
				$sigma_camp2 = $sigma_camp2 + $birdcount{$camp2}->{$bird_in_datasheet};
				#$birdcount{$camp1}->{$bird_in_datasheet};
			}
			$bray_curtis_coeff = $braycurtis_i/($sigma_camp1+$sigma_camp2)*100;
			#print "THE VAL: $bray_curtis_coeff\n";
			push @row_values, $bray_curtis_coeff;
		}
		@row_to_write = ("$camp1", @row_values);
		$worksheet12->write_row($row_number, 0, \@row_to_write);
	  $row_number++;
	}

	print color("green"), "                    [DONE]\n", color("reset");
}

if ($Bray_Curtis_from_Checklist) {
	# Sheet13: Bray Curtis from Checklist
	print "Sheet13: Bray Curtis from Checklist";;
	my $worksheet13 = $workbook->add_worksheet("BrayCurtis_Checklist");

	@row_to_write = ("Camp", @camps);
	$worksheet13->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $camp1 (@camps) {
		@row_values = ();
		foreach my $camp2 (@camps) {
			$checklist;
			$braycurtis_i = 0;
			$sigma_camp1 = 0;
			$sigma_camp2 = 0;
			$bray_curtis_coeff = 0;
			#print "\n\n\nCAMPS: $camp1 $camp2\n\n";
			foreach my $bird_in_datasheet (@birds_in_datasheet) {

				if (($birdcount{$camp1}->{$bird_in_datasheet} > 0) and ($birdcount{$camp2}->{$bird_in_datasheet} > 0)) {
					$checklist = 1;
				} else {
					$checklist = 0;
				}
				$braycurtis_i = $braycurtis_i + 2*$checklist;
				$sigma_camp1++, if ($birdcount{$camp1}->{$bird_in_datasheet} > 0);
				$sigma_camp2++, if ($birdcount{$camp2}->{$bird_in_datasheet} > 0);
				#print "Bird: $bird_in_datasheet\t C1: $birdcount{$camp1}->{$bird_in_datasheet}\t$birdcount{$camp2}->{$bird_in_datasheet}\tCheck: $checklist\t Sigma: $braycurtis_i\n";
				#$birdcount{$camp1}->{$bird_in_datasheet};
			}
			$bray_curtis_coeff = $braycurtis_i/($sigma_camp1+$sigma_camp2)*100;
			#print "THE VAL: $bray_curtis_coeff\n";
			push @row_values, $bray_curtis_coeff;
		}
		@row_to_write = ("$camp1", @row_values);
		$worksheet13->write_row($row_number, 0, \@row_to_write);
	  $row_number++;
	}

	print color("green"), "                    [DONE]\n", color("reset");
}

if ($Dips) {
	# Sheet14: Dips
	print "Sheet14: Dips";
	my $worksheet14 = $workbook->add_worksheet("Dips");

  # Compare with a more comprehensive list
  checkForDips();

	@row_to_write = ("No", "IOC No", "Dip");
	$worksheet14->write_row(0, 0, \@row_to_write);

	$row_number = 1;
	foreach my $bird_dip (@birds_dips) {

		@row_to_write = ($row_number, $birds{$bird_dip}->{ioc}, $bird_dip);
		$worksheet14->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	print color("green"), "                                          [DONE]\n", color("reset");
}

if ($Ebird) {
	# Sheet15: Ebird
	print "Sheet15: Ebird Checklist Format";
	my $worksheet15 = $workbook->add_worksheet("Ebird");

	my @fill;
	my @dates;
	my @notes;
	my @num_observers,;
	my $campxtransect = 0;

  foreach my $camp (@camps) {
    foreach my $transect (@{$transects{$camp}}) {
		  $campxtransect++;
      push (@camp_transects, "$camp"."_"."$transect");
      push (@dates, $date{"$camp"."_"."$transect"});
			push (@notes, "Observed by ".$observers{"$camp"."_"."$transect"}.". $surveyname");
			my @observers_list = split (",", $observers{"$camp"."_"."$transect"});
			push (@num_observers, $#observers_list+1);
      push (@camp_camp, "$areaname--$camp");
    }
  }

	#my $campxtransect = ($#camps+1) * ($#transects{$camp}+1);
	$row_number = 0;
	@row_to_write = ("", "", @camp_camp);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Latitude");
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Longitude");
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	#@row_to_write = ("Date", "", @camp_transects);
	@row_to_write = ("Date", "", @dates);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Start Time", "", @start_times);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@fill = ("TN") x $campxtransect;
	@row_to_write = ("State", "", @fill);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@fill = ("IN") x $campxtransect;
	@row_to_write = ("Country", "", @fill);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@fill = ("Traveling") x $campxtransect;
	@row_to_write = ("Protocol", "", @fill);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Num Observers", "", @num_observers);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Duration (min)", "", @durations);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@fill = ("Y") x $campxtransect;
	@row_to_write = ("All Obs Reported (Y/N)", "", @fill);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Dist Traveled (Miles)");
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Area Covered (Acres)");
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	@row_to_write = ("Notes", "", @notes);
	$worksheet15->write_row($row_number, 0, \@row_to_write);
	$row_number++;

	foreach my $bird_in_datasheet (@birds_in_datasheet) {

		@row_values = ();
		foreach my $camp_transect (@camp_transects) {
      if ($ebirdcount{$camp_transect}->{$bird_in_datasheet}) {
        push @row_values, $ebirdcount{$camp_transect}->{$bird_in_datasheet};
      } else {
        push @row_values, "";
      }
		} # camp

		if ($birds{$bird_in_datasheet}->{cmnnam_clm}) {
		  @row_to_write = ($birds{$bird_in_datasheet}->{cmnnam_clm}, "", @row_values);
	  } else {
		  @row_to_write = ($bird_in_datasheet, "", @row_values);
	  }

		$worksheet15->write_row($row_number, 0, \@row_to_write);
		$row_number++;
	}

	print color("green"), "                        [DONE]\n", color("reset");
}



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
			$birds{$row_contents[$lookup_cmnname_col-1]} = {ioc 		   => $row_contents[$lookup_ioc_col-1],
																											order 	   => $row_contents[$lookup_order_col-1],
																											family 	   => $row_contents[$lookup_family_col-1],
																											scinam 	   => $row_contents[$lookup_sciname_col-1],
																											cmnnam_clm => $row_contents[$lookup_cmnname_clm_col-1],
																											scinam_clm => $row_contents[$lookup_sciname_clm_col-1],
																											endemic	   => $row_contents[$lookup_endemic_col-1],
																											wg_endemic => $row_contents[$lookup_wg_endemic_col-1],
																											guild		   => $row_contents[$lookup_guild_col-1],
																											redlist	   => $row_contents[$lookup_redlist_col-1],
																											biome		   => $row_contents[$lookup_biome_col-1],
																											range		   => $row_contents[$lookup_range_col-1]};
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

			# Make an array of all India endemics
			unless (grep(/^$row_contents[$lookup_endemic_col-1]$/, @endemics)) {
				push @endemics, $row_contents[$lookup_endemic_col-1], if (($row_contents[$lookup_endemic_col-1]) and ($row_contents[$lookup_endemic_col-1] ne "-"));
			}

			# Make an array of all Western Ghats endemics
			unless (grep(/^$row_contents[$lookup_wg_endemic_col-1]$/, @wg_endemics)) {
				push @wg_endemics, $row_contents[$lookup_wg_endemic_col-1], if (($row_contents[$lookup_wg_endemic_col-1]) and ($row_contents[$lookup_wg_endemic_col-1] ne "-"));
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
	my $prev_transect = "none";
	my $prev_time = "00:00";

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
										date 			=> $row_contents[$datasheet_date_col-1],
										time 			=> $row_contents[$datasheet_time_col-1],
										bird 			=> $row_contents[$datasheet_bird_col-1],
										num 			=> $row_contents[$datasheet_num_col-1],
										db1 			=> $row_contents[$datasheet_db1_col-1],
										db2 			=> $row_contents[$datasheet_db2_col-1],
										db3 			=> $row_contents[$datasheet_db3_col-1],
										db4 			=> $row_contents[$datasheet_db4_col-1],
										habitat		=> $row_contents[$datasheet_habitat_col-1],
										observers	=> $row_contents[$datasheet_observers_col-1],
										remarks		=> $row_contents[$datasheet_remarks_col-1]};

			# Make an array of all birds seen atleast once in datasheet
			unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @all_birds_in_datasheet)) {
				push @all_birds_in_datasheet, $row_contents[$datasheet_bird_col-1];
			}

			# Print a warning if the entry in the datasheet has no match in the lookup
			unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @bird_names)) {
				push @birdname_nomatch, $row_contents[$datasheet_bird_col-1], unless (grep(/^$row_contents[$datasheet_bird_col-1]$/, @birdname_nomatch));
			}


			if ($row_contents[$datasheet_transect_col-1] ne $prev_transect) {
				push @start_times, $row_contents[$datasheet_time_col-1];
				push @end_times, $prev_time, unless ($prev_time eq "00:00");
				#print "\nTransect Change: $row_contents[$datasheet_camp_col-1] $row_contents[$datasheet_camp_col-1]  Now: $row_contents[$datasheet_time_col-1]  Last: $prev_time\n";
			}
			$prev_transect = $row_contents[$datasheet_transect_col-1];
			$prev_time = $row_contents[$datasheet_time_col-1];


			# Make an array of all camps
			unless (grep(/^$row_contents[$datasheet_camp_col-1]$/, @camps)) {
				push @camps, $row_contents[$datasheet_camp_col-1];
			}

			# Make an array of all transects
			unless (grep(/^$row_contents[$datasheet_transect_col-1]$/, @{$transects{$row_contents[$datasheet_camp_col-1]}})) {
				push @{$transects{$row_contents[$datasheet_camp_col-1]}}, $row_contents[$datasheet_transect_col-1];
			}

			# Assign Date and observers to camp_transect
			my $camp_transect_var = "$row_contents[$datasheet_camp_col-1]"."_"."$row_contents[$datasheet_transect_col-1]";
			$date{$camp_transect_var} = $row_contents[$datasheet_date_col-1];
			$observers{$camp_transect_var} = $row_contents[$datasheet_observers_col-1];

		} # Row
		push @end_times, $row_contents[$datasheet_time_col-1];

		foreach my $i (0 .. $#start_times) {
			my @hrmin_start = split (/[:\.]/, $start_times[$i]);
			my $hrminval_start = $hrmin_start[0]*60 + $hrmin_start[1];
			my @hrmin_end = split (/[:\.]/, $end_times[$i]);
			my $hrminval_end = $hrmin_end[0]*60 + $hrmin_end[1];
			#print "$start_times[$i] - $end_times[$i]  ", $hrminval_end-$hrminval_start, "\n";
			push @durations, $hrminval_end-$hrminval_start;
		}
		

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

		if ($Checklist) {
			push @birds_dips, $bird, unless (grep(/^$bird$/, @survey_checklist));
		} else {
			push @birds_dips, $bird, unless (grep(/^$bird$/, @birds_in_datasheet));
		}
	}
}




sub sumItAllUp {
 
	print "\n\nAnalysing Data";

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


		# SPECIES Count (Per camp_transect for EBIRD)
		foreach my $camp (@camps) {
		  my $camp_transect = "$camp"."_"."$$datum{transect}";
			$ebirdcount{$camp_transect}->{$$datum{bird}} = $ebirdcount{$camp_transect}->{$$datum{bird}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}


		# India ENDEMIC Count (Per camp and total)
		foreach my $camp (@camps) {
			$endemiccount{$camp}->{$birds{$$datum{bird}}->{endemic}} = $endemiccount{$camp}->{$birds{$$datum{bird}}->{endemic}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$endemiccount{Total}->{$birds{$$datum{bird}}->{endemic}} = $endemiccount{Total}->{$birds{$$datum{bird}}->{endemic}} + $$datum{num};

		# Western Ghats ENDEMIC Count (Per camp and total)
		foreach my $camp (@camps) {
			$wg_endemiccount{$camp}->{$birds{$$datum{bird}}->{wg_endemic}} = $wg_endemiccount{$camp}->{$birds{$$datum{bird}}->{wg_endemic}} + $$datum{num}, if ($$datum{camp} eq "$camp");
		}
		$wg_endemiccount{Total}->{$birds{$$datum{bird}}->{wg_endemic}} = $wg_endemiccount{Total}->{$birds{$$datum{bird}}->{wg_endemic}} + $$datum{num};

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

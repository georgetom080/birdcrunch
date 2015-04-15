# Utility to convert lat-long from various formats to DD.DDDDD

my @split_cord;
my $degree;
my $minute;
my $second;
my $dddddd;

open FILE, $ARGV[0] or die $!;
while (<FILE>) {

  chomp $_;
  @split_cord = split (" ", $_);

  # DD MM.MMMM
  if ($#split_cord == 1) {
    $degree = $split_cord[0];
    $minute = $split_cord[1];
    $dddddd = $degree + 1.0/60*$minute;

  #DD MM SS.SSSS
  } elsif ($#split_cord == 2) {
    $degree = $split_cord[0];
    $minute = $split_cord[1];
    $second = $split_cord[2];
    $dddddd = $degree + 1.0/60*$minute + 1.0/3600*$second;
  
  # DD.DDDD
  } elsif ($#split_cord == 0) {
    $dddddd = $split_cord[0];
  }

  # Print the result in DD.DDDDD format
  print "$dddddd\n";
}
close (FILE);

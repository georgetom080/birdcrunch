# Utility to compare 2 files line by line and print the differerntial

open (FILE1, $ARGV[0]);
while (<FILE1>) {
  chomp $_;
	push @file1, $_;
}
close (FILE1);

open (FILE2, $ARGV[1]);
while (<FILE2>) {
  chomp $_;
	push @file2, $_;
}
close (FILE2);

print "\n\nOnly in $ARGV[0]\n";
foreach my $file1 (@file1) {
  #print "$file1\n";
  print "$file1\n", unless (grep(/$file1/, @file2));
}

print "\n\nOnly in $ARGV[1]\n";
foreach my $file2 (@file2) {
  #print "$file2\n";
  print "$file2\n", unless (grep(/$file2/, @file1));
}

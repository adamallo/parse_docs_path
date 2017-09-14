use strict;
use warnings;
use File::Basename;

##config
my $timesinfo=9.0;
our $DEBUG=0;

my $usage="\nUsage: perl $0 root_directory outputname\n----------------------------------------------------------\n\nThis script parses Microsoft Word files with reports in the format of Northwest Zoopath, obtaining their case numbers and history information for posterior parsing. It generates three output files, outputname.tsv with the output tabular database, outputname.errors with a list of files for which the number of detected cases is 0, and outputname.warnings with a list of files in which the number of parsed cases is not the same as the number of the words \"case\" or \"history\". Files that generate warnings may containg unparsed cases due to format/parsing issues However, the number of false positives is expected to be very high since the word history can be very common.\n";

my $folder=".";
my $outfile="outputfile";

if (scalar @ARGV != 2 || !-d $ARGV[0] )
{
	print("\nYour input options are not correct\n");
	print $usage;
	exit 1;
} 
else {
	
	$folder = $ARGV[0];
	$outfile = $ARGV[1];
}

my @filesdoc=`find $folder -type f -name '*.[dD][oO][cC]'`;
my @filesdocx=`find $folder -type f -name '*.[dD][oO][cC][xX]'`;
my @files=(@filesdoc,@filesdocx);
my $nfiles= scalar @files;

my $period=int($nfiles/$timesinfo+0.5); #No negative numbers, otherwise this will fail
my $n_noninfo=0;

print("Parsing $nfiles files\n");
my $file="";
my @suffixes=(".doc",".docx");
my $content;
my $backup=$/;
$/="";
my $filename;
#my $nsubs;
my $nparsed;
my @parsed;
my $nhistories;
my $ncases;
my $id="";
my $history="";

open(my $OUTPUT,">$outfile.tsv") or die "Impossible to write the tsv output file";
print($OUTPUT "caseid\thistory\n");
open(my $OUTPUTWARNING, ">$outfile.warnings") or die "Impossible to write the warning output file";
open(my $OUTPUTERROR, ">$outfile.errors") or die "Impossible to write the error output file";

for (my $i=0; $i<$nfiles; ++$i)
{	
	$n_noninfo=$n_noninfo+1;
	if ($n_noninfo==$period)
	{
		my $percent=$i/$nfiles*100;
		printf("Parsed the $percent percent of the files, be patient!\n");
		$n_noninfo=0;
	}
	$file=$files[$i];
	$file=~s/ /\\ /g;
	#$filename=fileparse("$file",@suffixes);
	chomp($file);
	$content=`textutil -stdout -convert txt $file`;
	#$content=`java.exe -jar tika-app-1.15.jar -t $file` ##windows cygwin
	$content=~s/\n/ /msg;
	$ncases=()=$content=~m/Case/msig; ##Matches in list context to get the different matches and then scalar context to get the amount of them we have
	$nhistories=()=$content=~m/HISTORY/msig;
	#@parsed=$content=~m/^.*Case\ No\.:\t[^\s]*.*? HISTORY.*?CLINICAL/msg;
	
	@parsed=$content=~m/Case\ No\.:\t[^\s]*.*? HISTORY.*?CLINICAL/msg;
	#@parsed=$content=~m/Case\ No\.:[\t\s][^\s]*.*? HISTORY.*?CLINICAL/msg; #Windows
	$nparsed=scalar @parsed;
	#$nsubs=$content=~s/^.*Case\ No\.:\t([^\s]*).* HISTORY(.*)CLINICAL.*$/$1\n$2/ms;

	#if ($nsubs == 0)
	if ($nparsed == 0)
	{
		print("\tERROR: No entries parsed in the file $file\n");
		print($OUTPUTERROR "$file\n");
	}
	else
	{
		if ( $nhistories != $ncases || scalar $nparsed != $nhistories)
 		{
			print("\tWARNING: the number of parsed entries $nparsed,and the words \"history\": $nhistories and \"case\": $ncases are not the same in the file $file\n");
			print($OUTPUTWARNING "$file\n");
		}
		
		if($DEBUG)
		{
			print("\tDEBUG: $nparsed entries detected in the file $file\n");
		}
		for my $match (@parsed)
		{
			$match=~s/Case\ No\.:\t([^\s]*).* HISTORY(.*)CLINICAL/$1\n$2/ms;	
#			$match=~s/Case\ No\.:[\t\s]([^\s]*).* HISTORY(.*)CLINICAL/$1\n$2/ms; #Windows

			($id,$match)=split("\n",$match);
			$match=~s/\t/    /g;
			print($OUTPUT "$id\t$match\n");
		}
	}
}
close($OUTPUT);
close($OUTPUTWARNING);
close($OUTPUTERROR);

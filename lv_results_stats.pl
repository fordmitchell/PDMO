#!/usr/intel/bin/perl5.14.1 -w
#### LV Summary Extractor
#### Updated 7/9/14 - bspaynic - Added single cell selection and flip option
#### Updated 7/9/14 - bspaynic - Checks to see if -cell exists
#### Updated 7/10/14 - bspaynic - Added ability to mail results to other users
#### Updated 10/17/14 - bspaynic - added singleflow extraction capability

BEGIN {
	push(@INC, "/p/alliance/cad/nova/PerlMod");
	push(@INC, "/p/hiproot/pdmo/repo/01/release/common/perl/layout_scripts/lv_results/modules");
    our $start_run = time(); 
}
    use strict;
    use warnings;
    use Data::Dumper qw(Dumper);
    use Spreadsheet::WriteExcel;
	use Excel::Writer::XLSX;
	use Cwd;
	use MIME::Lite;
	use Getopt::Long;	
	use Pod::Usage;
	use Env; 
	my $homedir = $ENV{'HOME'};
	my $site = $ENV{'EC_SITE'};
	my $item;
	my $end_run = time();
	my $run_time = ($end_run - our $start_run)/1000;
	no warnings 'uninitialized';

	my $start_date = `date +'%m/%d/%Y'`;
	my $start_time = `date +'%T'`;


#### Get list of stats files and get uniq and sorted list
   	 
    my @files;
    my $cellopt;
    my $singleflow;
    my $flip;
    my $email = '';
    my $help;
    my $list;
    my $outtype; 
    my $listitem;
    my @file;
    my @merged;
    my @collect;
    my $all;
    my $nolarge;
    my $dir;

    GetOptions ('cell=s' 		=> \$cellopt,
				'list=s' 		=> \$list, 
				'flip'  		=> \$flip,
				'email=s' 		=> \$email,
				'all'     		=> \$all,
				'nolarge' 		=> \$nolarge,
				'singleflow=s'  => \$singleflow,
				'help'   		=> \$help); 

### Handle ICVLVS stats names
        my @icvfiles = glob("*.icvlvs.stats");

        foreach my $nm (@icvfiles){
                my @icvsep = split('\.', $nm);
                my $a = $icvsep[0];
                my $b = $icvsep[1];
                my $c = $icvsep[2];
                my $nw = $a . "\." . $b . "_" . $c . "\.stats";
                rename $nm, $nw;

        }

       if(defined $cellopt){
		my $dir = '.';
		my @fcnt = <$dir/$cellopt.*.stats>;
		my $fcnt = @fcnt;

		if($fcnt > 0){ 
	                print "\nExtracting cell: $cellopt \n";
			@files = glob("$cellopt.*.stats");
			$outtype = "Extracting cell(s):  $cellopt";
		}
		else {
		print "\nCellname $cellopt stats files not found\n";
		print "Exiting...\n\n";
		exit;
		}
	}
 
        if(defined $all){ 
		my $dir = '.';
                my @fcnt = <*.stats>;
                my $fcnt = @fcnt;
                if($fcnt > 0){
                print "\nExtracting cells in directory \n";
		$outtype = "Extracting entire directory";
                @files = glob("*.stats");

                }
	}

        if(defined $list){

        open(my $data, '<', $list) or die "Could not open '$list' $!\n";
                        while (my $line = <$data>) {
                           chomp $line;
                           push @file, $line;
                        }

                foreach $listitem (@file){
                        @collect = glob("$listitem.*.stats");
                        @merged = (@merged, @collect);
                }

        @files = @merged;
	print "Extracted $list of cells\n";
	}

	if(defined $help){
		print "\n\n LV PDS LOG Extractor\n";
		print "\n Usage: lvsresultssum.pl [options]\n\n";
		print "   -all    extracts all data in the current directory\n";
		print "   -singleflow extracts singleflow (use with -all or -cell)\n";
		print "   -cell <cell>  extract data from specific cell\n";
		print "   -list <list>  provide a list of cells to extract\n"; 
		print "   -flip   flip the column/row headings\n";
		print "   -nolarge don't include large flows like hip_den_reuse, cmden_collat\n";
		print "   -email  CC other users on results\n";
		print "   -help   help information\n\n";
		print " SUMMARY\n";
		print "\n This script extracts all the stats files \n";
		print " in a directory. All the error counts are then\n";
		print " tabulated and send to you in an email.\n\n";
		exit;
	}

    my $cwd = getcwd;
    my @cell;
    my @flow;

    foreach (@files) {
	my @separated = split('\.', $_);
	push @cell, $separated[0];
	push @flow, $separated[1];   
    }

my %seen;
my @uniq_cell;
foreach $item (@cell) {
    unless ($seen{$item}) {
        # if we get here, we have not seen it before
        $seen{$item} = 1;
        push(@uniq_cell, $item);
    }
}

if(@uniq_cell){
} else {	print "The script failed, either because\n";
		print "\n\nYou must use an option... either -all, -list, -cell or -help\n";
		print "Or you provide a list with no valid cellnames\n";
                print "\n\n LV PDS LOG Extractor\n";
                print "\n Usage: lvsresultssum.pl [options]\n\n";
                print "   -all    extracts all data in the current directory\n";
                print "   -cell <cell>  extract data from specific cell\n";
                print "   -list <list>  provide a list of cells to extract";
                print "   -flip   flip the column/row headings\n";
                print "   -email  CC other users on results\n";
                print "   -help   help information\n\n";
                print " SUMMARY\n";
                print "\n This script extracts all the stats files \n";
                print " in a directory. All the error counts are then\n";
                print " tabulated and send to you in an email.\n\n";
                exit;

}

my @uniq_flow;
foreach $item (@flow) {
    unless ($seen{$item}) {
        # if we get here, we have not seen it before
        $seen{$item} = 1;
        push(@uniq_flow, $item);
    }
}

if ($singleflow){
	splice(@uniq_flow);
	@uniq_flow = $singleflow;	
}

if ($nolarge){
    @uniq_flow = grep ! /hip_den_reuse/,@uniq_flow;
    @uniq_flow = grep ! /cmden_collat/, @uniq_flow;
    @uniq_flow = grep ! /hip_den_bronze/, @uniq_flow;
}

 my $outfile = "$homedir/lv_summary.xlsx";
 my $outfilename =  "lv_summary.xlsx";

 my $workbook = Excel::Writer::XLSX->new("$outfile");
 my $worksheet1 = $workbook->add_worksheet("Results");

#### Build formats for Excel

 my $bold = $workbook->add_format();
 $bold->set_bold();

 my $dirty = $workbook->add_format();
    $dirty->set_color('red');
    $dirty->set_align('center');

 my $clean = $workbook->add_format();
    $clean->set_color('blue');
    $clean->set_align('center');
 my $cleantxt = "clean";

 my $nodata = $workbook->add_format();
    $nodata->set_color('black');
    $nodata->set_bg_color('yellow');
    $nodata->set_align('center');

#### Build headers for columns and rows 

if($flip){

      my $count=1;
        my $cel;

        foreach $cel (@uniq_flow){
        my $format = $workbook->add_format();
        $format->set_rotation( 90 );
        $format->set_align('center');
        $format->set_bold();
        $worksheet1->write(0, $count, $cel, $format);
        $count++;
        }

        $count=1;
        my $row;
        foreach $row (@uniq_cell){
        $worksheet1->write($count, 0, $row, $bold);
        $count++;
        }

} else {

        my $count=1;
        my $cel;

        foreach $cel (@uniq_cell){
        my $format = $workbook->add_format();
        $format->set_rotation( 90 );
        $format->set_align('center');
        $format->set_bold();
        $worksheet1->write(0, $count, $cel, $format);
        $count++;
        }

        $count=1;
        my $row;
        foreach $row (@uniq_flow){
        $worksheet1->write($count, 0, $row, $bold);
        $count++;
        }

}
	$worksheet1->set_column('A:A',25);
	$worksheet1->set_row('0',200);

	my $c;
	my $f;
	my $rowcnt=1;
	my $colcnt=1;
	my $cr = chr(13);
	my $tab = chr(9);

#### Build Data Table
	
foreach $c (@uniq_cell){
		
	foreach $f (@uniq_flow) {
	print "Extracting results for $c, $f\n"; 
	my $cf = $c . "\." . $f . "\.stats";	

	if(-e $cf){ 
			open(DAT, $cf);
			my @a=<DAT>;			
			close(DAT);
			my @l = grep /Total Errors/, @a;
					
		if(@l) {
			$l[0] =~ s/Total Errors =//g;
			my $num = $l[0] + 0;
			
		if($flip){
                        if($num == 0){
                            $worksheet1->write($colcnt,$rowcnt,'clean',$clean);
			} else {
			    $worksheet1->write($colcnt,$rowcnt,$num,$dirty);
			}
		} else {
			if($num == 0){
                          $worksheet1->write($rowcnt,$colcnt,'clean',$clean);
			} else {	
			  $worksheet1->write($rowcnt,$colcnt,$num,$dirty);
			}
		}	

			my @statfile = glob("$c.$f.stats");	
			my $statfile = join '', @statfile;

			my @com; 
      			open(my $data, '<', $statfile) or die "Could not open '$statfile' $!\n";
			my $linecnt;
			my $i=1;
			
			my @lines = map scalar(<$data>), 1..500;
			close($data);

			foreach my $ln (@lines){ 
			$ln =~s/\t/\ - /g;
  			push @com, $ln ;
			}

			$linecnt++;	
			my $comment = join '', @com;
			my $height = $linecnt * 7;

			if($flip){
				$worksheet1->write_comment($colcnt,$rowcnt, $comment, width => 700, height => 500, font => 'Courier');
			} else{
				$worksheet1->write_comment($rowcnt,$colcnt, $comment, width => 700, height => 500, font => 'Courier');
			}
			$linecnt = 0;
		} 	
		else
		{
				if($flip){
				$worksheet1->write($colcnt,$rowcnt,'clean', $clean);
				} else
				{
				$worksheet1->write($rowcnt,$colcnt,'clean', $clean);
				}
		}	
	} else
		{
                if($flip){
                $worksheet1->write($colcnt,$rowcnt,'no data', $nodata);
                } else
                {
                $worksheet1->write($rowcnt,$colcnt,'no data', $nodata);
                }
		}

	$rowcnt++;

	}
$colcnt++;
$rowcnt=1;
}

$workbook->close(); 

#### Mail Results to User 
if(! -e $outfile)
 {       print "-ERROR-sendmail-excel- $outfile not exist\n";
         exit;
 }
		my $DomainAddress ="-";
		my $user=`/usr/bin/whoami`;
        chomp $user;
        my $detail= `/usr/intel/bin/cdislookup -i $user`;
        if ( $detail =~ /\sDomainAddress\s*=\s*(.*)\sEmptype/ ){
			if ( $1 ne '' ){
			$DomainAddress = $1;
			}
        }
        my $msg = MIME::Lite->new(From		=> 'no-reply@intel.com',
								To			=> $DomainAddress,
								Cc			=> $email, 
								Bcc			=> '',
								Subject		=> "lv pds/logs extractor results",
								Type		=> 'multipart/mixed')
								or die "PROBLEM opening MIME object: $!";
    $msg->attach(
			Type     	=> 'TEXT',
			Data     	=> "LV PDS/LOGS results for $cwd\nSubmitted by: $user\n$outtype"
    	);
        $msg->attach (
           Type			=> 'image/png',
           Path			=> $outfile,
           Filename		=> $outfilename,
           Disposition	=> 'attachment'
        ) or die "Error adding $outfile : $!\n";

$msg->send() or die "PROBLEM sending MIME mail: $!";

my $cur_date = `date +'%m/%d/%Y'`;
my $cur_time = `date +'%T'`;
my $datestring = localtime();

 system("/usr/intel/bin/mysql -h maria3050-us-fm-in.icloud.intel.com -umig_pde_anal_rw -pvEkJ0Mz16Z647Nl mig_pde_analytics2 -e \"INSERT INTO migpde_script_invocations (script,user,site,description,dateval,timeval,timestart,timeend,runtype) VALUES ('migLVstats.pl', '$user', '$site', 'summarize lv results from stats', '$start_date','$start_time', '$start_time', '$cur_time', 'dev')\"");

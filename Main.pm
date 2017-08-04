#!/usr/bin/perl
package EA::Main;

use strict;
use warnings;

use FindBin;
use lib "$FindBin::RealBin/lib/";

use Win32::OLE;
use Win32::ODBC;
use Win32::OLE::Const 'Microsoft DAO';

use Data::Dumper;
use File::Copy::Recursive qw(dircopy);

use constant DEBUG => 0;

use vars qw( $VERSION $DEBUG );

BEGIN {
    $VERSION = '0.01';
    $DEBUG   = 0;
}

my @PublicMethods = qw/ GetAllLatest ExportHTML close /;

my $MaxRecursionDepth = 666;
my $ExportTemplates = "export";

sub new {
	my ($proto, %args) = @_;
    my $class = ref($proto) || $proto;
				
	my $file = $args{'file'};
				
    my $self = {
		file 	=> $args{'file'},
		output 	=> $args{'output'},
		html_template => $args{'html_template'}
	};
	
	printf(" \r\n");
	printf("=============================================\r\n");
	printf("Opening Repository '%s'\r\n", $file);
	
	$self->{MODEL} = Win32::OLE->new('EA.Repository', \&_OleQuit)
				  || Win32::OLE->GetActiveObject('EA.App') 
				  or die "Error initializing EA-Repository: $!\n";
	$self->{MODEL}->OpenFile($file) or die "Error at openin '$file: $!\n";
	
	printf("Repository opened\r\n");
	printf("=============================================\r\n");
	printf(" \r\n");
	
	bless $self, $class;

    return $self;
}

sub close {
	my $self = shift;
	
	$self->{MODEL}->CloseFile();
	
	printf(" \r\n");
	printf("=============================================\r\n");
	printf("Repository '%s' closed\r\n", $self->{'file'});
	printf("=============================================\r\n");
	printf(" \r\n");
}

sub GetAllLatest {
	my $self = shift;
	
	# Root-Package
	my $package = $self->{MODEL}->Models->GetAt(0) or die "Error at openin Root-Model: $!\n";
	
	printf(" \r\n");
	printf("=============================================\r\n");
	printf("'GetAllLatest' Package '%s' started \r\n", $package->Name);		
	printf(" \r\n");
	
	_GetLatestSubPackages('', 0, $package);
	
	printf(" \r\n");
	printf("'GetAllLatest' Package finished \r\n");
	printf("=============================================\r\n");
	printf(" \r\n");
}

sub ExportHTML {
	my $self = shift;	
	
	my $package = $self->{MODEL}->Models->GetAt(0) or die "Error at openin Root-Model: $!\n";
	my $projectInterface = $self->{MODEL}->GetProjectInterface() or die "Error getting Project Interface: $!\n";
	
	printf(" \r\n");
	printf("=============================================\r\n");
	printf("'Export HTML' Package '%s' started \r\n", $package->Name);
	
	$projectInterface->RunHTMLReport($package->PackageGUID, $self->{'output'}, "PNG", $self->{'html_template'}, ".html");
	dircopy($ExportTemplates, $self->{'output'});
	
	printf("'Export HTML' Package finished \r\n");
	printf("=============================================\r\n");
	printf(" \r\n");
}

sub COMPACT {
	my $self = shift;
	my ($file) = @_;

	my $DSN      = 'T-Bonds';
	my $Driver   = 'Microsoft Access Driver (*.mdb, *.accdb)';
	my $Desc     = 'US T-Bond Quotes';
	my $Dir      = 'D:\\';
	my $File     = 'Test.accdb';
	my $Fullname = "$Dir\\$File";
	my $File2	 = 'Test2.accdb';
	my $Fullname2 = "$Dir\\$File2";
	 
	# Remove old database and dataset name
	#unlink $Fullname if -f $Fullname;
	#Win32::ODBC::ConfigDSN(ODBC_REMOVE_DSN, $Driver, "DSN=$DSN")
	#					   if Win32::ODBC::DataSources($DSN);
	 
	# Create new database
	my $dao = Win32::OLE->new('DAO.DBEngine.35', 'Quit') or die "$!";	
	

	$dao->CompactDatabase($Fullname, $Fullname2) or die "something wrong?";

	
	 
	# Add new database name
	#Win32::ODBC::ConfigDSN(ODBC_ADD_DSN, $Driver,
#			"DSN=$DSN", "Description=$Desc", "DBQ=$Fullname",
#			"DEFAULTDIR=$Dir", "UID=", "PWD=");
	

	#my $dbe = Win32::OLE->CreateObject('DAO.DbEngine.36') or die "Error creating DAO.DBEngine object" . Win32::OLE::DumpError();

	#print Dumper($dbe);

	#my $dbh = DBI->connect('dbi:ADO:Provider=Microsoft.Jet.OLEDB.12.0;Data Source='.$file) or die $DBI::errstr;
	
	#my $dbh = DBI->connect('dbi:ODBC:driver=Microsoft Access Driver (*.mdb, *.accdb);dbq='.$file);
	
	#print Dumper($dbh);	
		
	#my $handler = Win32::OLE->new('Microsoft.ACE.OLEDB.12.0') or die "geht net $!\r\n";
	
	# my $dsn	= "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$file;";
	
	
	# print "Connection String = '$dsn'\r\n";
	
	# my $handler = Win32::ODBC->new($dsn);
	
		# Win32::ODBC::DumpError();
	# if (undef $handler) {
	# } else {
		# print "Con = '" . $handler->Connection() . "'\r\n";
	# }
	
	# my $dbe = Win32::OLE->GetActiveObject('Access.Application') 
	#	    || Win32::OLE->new('Access.Application', 'Quit')
#		    or die "Error creating/opening Access-Instance";
		   
	 #print "Access App handle:",$dbe,"\n";
	
	# $dbe->{Visible} = 1;
		   
	 #my $dbs = $dbe->DBEngine->OpenDatabase($file) or die "Error opening database '$file' $!\n";
	
	# my $dbs = $dbe->OpenDatabase($file) or die "Error openening database '$file': $!\n";
	# $dbs->Version or die "Error getting database information: $!\n";
	# print "Compacting Database '$dbs->{Name}', '$dbs->{Version}'\n";
	# $dbs->Close;
	
	#$dbe->CompactDatabase($file, "D:\test2.accdb");

	# $dbs = $dbe->OpenDatabase($file) or die "Error openening database '$file': $!\n";
	# $dbs->Version or die "Error getting database information: $!\n";
	# print "Compacted Database '$dbs->{Name}', '$dbs->{Version}'\n";
	# $dbs->Close;	
}

#### Internal Methods

# Update Packages
sub _GetLatestSubPackages {

	my ($identation, $recursionLevel, $currentPackage) = @_;

	if ($recursionLevel < $MaxRecursionDepth) {
						
		my $count = $currentPackage->Packages->Count;
		for(my $i = 0; $i < $count; $i++) {

			my $childPackage = $currentPackage->Packages->GetAt($i);
			if (index($childPackage->Name, "subdomain requirements") != -1) {
				printf("%s Skipping %s (PackageId='%s')\r\n", $identation, $childPackage->Name, $currentPackage->PackageId);
			} else {
				_GetLatestSubPackages($identation . "   ", $recursionLevel++, $childPackage);
			}			
		}
	}
	
	if ($currentPackage->isVersionControlled) {
		printf("%s Updating %s (PackageId='%s')\r\n", $identation, $currentPackage->Name, $currentPackage->PackageId);
		$currentPackage->VersionControlGetLatest(0);
	} else {
		printf("%s Unversioned %s (PackageId='%s')\r\n", $identation, $currentPackage->Name, $currentPackage->PackageId);
	}
}

# CleanUp EA
sub _OleQuit {
	my $self = shift;
	$self->Exit();
}

1;
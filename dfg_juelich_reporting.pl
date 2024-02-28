#!/usr/bin/perl -w

#######################################################################
#               Exportiert Publikations- und Kostendaten              #
#              Reporting fuer das Forschungszentrum Jülich            #
#                        Reportingjahr: 2022                          #
#                     Author: Volker Schallehn                         #
#                        Version 16.02.2023                           #
#######################################################################

use FindBin;
use lib "$FindBin::Bin/../perl_lib";
use Excel::Writer::XLSX;
use utf8;

use EPrints;
use strict;

#######################################################
my $datum = "2022";
#######################################################


my $repository_id = 'costs';

my $eprints = EPrints->new;
my $session = $eprints->repository( $repository_id );
exit( 0 ) unless( defined $session );

my $ds_archive = $session->dataset( "archive" );
if( !defined $ds_archive )
{
	print STDERR "Unbekannte Archiv-ID: $ds_archive\n";
	$session->terminate;
	exit 1;
}
my $ds_buffer = $session->dataset( "buffer" );
if( !defined $ds_buffer )
{
	print STDERR "Unbekannte Archiv-ID: $ds_buffer\n";
	$session->terminate;
	exit 1;
}

my $file = "reporting_".$datum.".xlsx";
my $workbook = Excel::Writer::XLSX->new( $file );
my $worksheet = $workbook->add_worksheet( 'mit DOI' );

# currency_format = workbook.add_format({'num_format': '[$$-409]#,##0.00'})

$worksheet->set_landscape();
$worksheet->set_header('LMU Open Access Fonds - Reporting '.$datum );
$worksheet->set_footer( '&CSeite &P von &N&RErstellt am &D.&M.&Y' );
$worksheet->hide_gridlines( 0 );
$worksheet->repeat_columns( 'A:A' );
$worksheet->print_row_col_headers();

my $header = $workbook->add_format( bold => 1, color => 'white', bg_color => 'green' );
$header->set_align( 'center' );
$header->set_border(1);

#my $costs = $workbook->add_format( bold => 1, color => 'white', bg_color => 'orange' );
#$costs->set_align( 'center' );
#$costs->set_border(1);

my $cells = $workbook->add_format( bold => 0, );
$cells->set_border(0);

my $cells_center = $workbook->add_format( bold => 0, );
$cells_center->set_align('center');

# Definition der Spaltenbreiten
$worksheet->set_column( 'B:B', 8 );
$worksheet->set_column( 'c:c', 30 );
$worksheet->set_column( 'D:D', 30 );
$worksheet->set_column( 'E:E', 15 );
$worksheet->set_column( 'F:F', 15 );
$worksheet->set_column( 'G:G', 25 );
$worksheet->set_column( 'H:H', 35 );
$worksheet->set_column( 'I:I', 15 );
$worksheet->set_column( 'J:J', 35 );
$worksheet->set_column( 'K:K', 30 );
$worksheet->set_column( 'L:L', 30 );
$worksheet->set_column( 'M:M', 30 );
$worksheet->set_column( 'N:N', 18 );
$worksheet->set_column( 'O:O', 18 );
$worksheet->set_column( 'P:P', 22 );
$worksheet->set_column( 'Q:Q', 20 );
$worksheet->set_column( 'R:R', 20 );
$worksheet->set_column( 'S:S', 50 );
$worksheet->set_column( 'Y:Y', 35 );
$worksheet->set_column( 'Z:Z', 20 );


$worksheet->write( 'A1', 'Kontrollfeld', $header );
$worksheet->write( 'B1', 'Eprintid', $header );
$worksheet->write( 'C1', 'DOI', $header );
$worksheet->write( 'D1', 'Name des Verlages', $header );
$worksheet->write( 'E1', 'Publikationsform', $header );
$worksheet->write( 'F1', 'CC-Lizenz', $header );
$worksheet->write( 'G1', 'Originalwährung', $header );
$worksheet->write( 'H1', 'Rechnungsbetrag in Originalwährung', $header );
$worksheet->write( 'I1', 'Euro netto', $header );
$worksheet->write( 'J1', 'Steuersatz', $header );
$worksheet->write( 'K1', 'Euro brutto', $header );
$worksheet->write( 'L1', 'Zuschussbetrag DFG', $header );
$worksheet->write( 'M1', 'nicht förderfähige Gebührenart', $header );
$worksheet->write( 'N1', 'Zuordnung zu Mitgliedschaft', $header );
$worksheet->write( 'O1', 'Zuordnung zu Transformationsvertrag', $header );
$worksheet->write( 'P1', 'Förderjahr', $header );
$worksheet->write( 'Q1', 'Rechnungsjahr', $header );
$worksheet->write( 'R1', 'Projektnummer/Projekt ID DFG', $header );
$worksheet->write( 'S1', 'DFG-Wissenschaftsbereich', $header );


my $results_archive = $ds_archive->search( filters => [ { meta_fields => [qw( eprintid )], value => "1-" }]);
my $results_buffer = $ds_buffer->search( filters => [ { meta_fields => [qw( eprintid )], value => "1-" }]);

my @kontrollfelder;
my @eprintids;
my @dois;
my @dates;
my @publishers;
my @types;
my @licenses;
my @currencies;
my @ammount;
my @oa_apc_values;
my @oa_apc_types;
my @oa_apc_net_values;
my @oa_apc_tax_rates;
my @oa_fundings;
my @payment_amounts;
my @euro_netto;
my @tax;
my @bruttos;
my @foerderbetrag;
my @gebuehrenart;
my @mitgliedschaft;
my @transformation;
my @funding_year;
my @billing_year;
my @dfg_number;
my @dfg_subjects;

my $net = 1;
my $tax = 1;


##### Archive
$results_archive->map (sub
{
	my( $session, $dataset, $eprint ) = @_;
	
	my $type = $eprint->get_value( 'type' );
	my $oa_funding = $eprint->get_value( 'oa_funding' );

		#### Kontrollfeld
		my $kontrollfeld = '=WENN(ANZAHL2(B2:K2; N2:P2; R2)<14;"Pflichtfeld fehlt!";"ok")';
		push (@kontrollfelder, $kontrollfeld);

		#### eprintids
		if( $eprint->exists_and_set( "eprintid" ) )
		{		
			my $eprintid = $eprint->get_value( 'eprintid' );
			push (@eprintids, $eprintid);
		}		
		
		#### DOI
		if( $eprint->exists_and_set( "id_number" ) )
		{		
			my $doi = $eprint->get_value( 'id_number' );
			# $doi =~ s/^(.*)$/https:\/\/doi.org\/$1/;
			$doi =~ s/doi://;
			push (@dois, $doi);
		}
		else {
			push (@dois, "" );
		}

		#### Verlag
		if( $eprint->exists_and_set( "publisher" ) )
		{			
		my $publisher = $eprint->get_value( 'publisher' );
		
		push (@publishers, $publisher);
		}
		else {
			push (@publishers, "" );
		}		

		#### Publikationsform
		if( $eprint->exists_and_set( "type" ) )
		{			
			my $type = $eprint->get_value( 'type' );
			$type =~ s/article.*/journal article/;
			push (@types, $type);
		}
		else {
			push (@types, "" );
		}	

	
		#### CC-Lizenz
		my @documents = $eprint->get_all_documents();
		foreach my $doc ( @documents )
		{
			if( $doc->is_public )
			{
				if( $doc->exists_and_set( "license" ) )
				{
					my $license = $doc->get_value ('license');
					
					$license =~ s/cc_by_4/CC BY/;
					$license =~ s/cc_by_sa_4/CC BY-SA/;
					$license =~ s/cc_by_nd_4/CC BY-ND/;
					$license =~ s/cc_by_nc_4/CC BY-NC/;
					$license =~ s/cc_by_nc_sa_4/CC BY-NC-SA/;
					$license =~ s/cc_by_nc_nd_4/CC BY-NC-ND/;
					
					push (@licenses, $license);
				}
				else {
					push (@licenses, "" );
				}
			}
		}
		
		
		
		#### Waehrung
		if( $eprint->exists_and_set( "oa_apc_type" ) )
		{
			my $currency = $eprint->get_value( 'oa_apc_type' );
			push (@currencies, $currency);
		}
		else {
			push (@currencies, "" );
		}

		#### Rechnungsbetrag
		if( $eprint->exists_and_set( "oa_apc_value" ) )
		{		
			my $oa_apc_value = $eprint->get_value( 'oa_apc_value' );
			$oa_apc_value =~ s/\./,/;
			push (@oa_apc_values, $oa_apc_value);
		}
		else {
			push (@oa_apc_values, "" );
		}
		
		#### EURO Nettobetrag
		if( $eprint->exists_and_set( "oa_apc_net_value" ) )
		{		
			my $oa_apc_net_value = $eprint->get_value( 'oa_apc_net_value' );
			my $oa_apc_net_type = $eprint->get_value( 'oa_apc_net_type' );
			$oa_apc_net_value =~ s/\./,/;
			push (@oa_apc_net_values, $oa_apc_net_value);
			
			$net = $net + 1;
			$tax = $tax + 1;
			# =I20+(I20*J20)
			# Muss in der Endversion geändert werden auf
			# =H20+(H20*I20)
			# da die Spalte 'eprintid' entfällt
			push (@bruttos, "=I".$net."+(I".$net."*J".$tax.")");
		}
		else {
			push (@oa_apc_net_values, "" );
		}
		
		

		#### Steuersatz
		if( $eprint->exists_and_set( "oa_apc_tax_rate" ) )
		{		
			my $oa_apc_tax_rate = $eprint->get_value( 'oa_apc_tax_rate' );
			$oa_apc_tax_rate =~ s/\./,/;
			$oa_apc_tax_rate /= 100;
			push (@oa_apc_tax_rates, $oa_apc_tax_rate);
		}
		else {
			push (@oa_apc_net_values, "" );
		}

		### Foerdersumme
		if( $eprint->exists_and_set( "payment_amount" ) )
		{
			my $payment_amount =  $eprint->get_value( 'payment_amount' );
			push (@payment_amounts, $payment_amount );
		}
		else {
			push (@payment_amounts, "" );
		}

		# Transformationsvertraege
		if( $eprint->exists_and_set( "transformation" ) )
		{
		my $transformation_contract = $eprint->get_value( 'transformation' );
		$transformation_contract =~ s/no/kein/;
		$transformation_contract =~ s/springer_deal/Springer (DEAL)/;
		$transformation_contract =~ s/wiley_deal/Wiley (DEAL)/;
		push ( @transformation, $transformation_contract );
		}
		else {
			push (@transformation, "kein" );
		}
		
		# Foerderjahr
		if( $eprint->exists_and_set( "date" ) )
		{
			my $date = $eprint->get_value( 'date' );
			$date =~ s/^(\d{4}).*$/$1/;
			push ( @funding_year, $date );
		}		
		else {
			push (@funding_year, "" );
		}
	
		
		
		### DFG-Projektnummern
		if( $eprint->exists_and_set( "oa_funding_dfg_id" ) )
		{		
			my @dfg_id =  @{$eprint->get_value( 'oa_funding_dfg_id' )};
			push ( @dfg_number, $dfg_id[0] );
		}
		else {
			push (@dfg_number, "" );
		}

		### DFG-Wissenschaftsbereiche
		if( $eprint->exists_and_set( "dfg_subjects" ) )
		{
			my @dfg_subject =  @{$eprint->get_value( 'dfg_subjects' )};
		
			if ( $dfg_subject[0] eq 'dfg01' )
			{
				$dfg_subject[0] = 'Geistes- und Sozialwissenschaften';
			}
			if ( $dfg_subject[0] eq 'dfg02' )
			{
				$dfg_subject[0] = 'Lebenswissenschaften';
			}			
			if ( $dfg_subject[0] eq 'dfg03' )
			{
				$dfg_subject[0] = 'Naturwissenschaften';
			}	
			if ( $dfg_subject[0] eq 'dfg04' )
			{
				$dfg_subject[0] = 'Ingenieurwissenschaften';
			}	
			if ( $dfg_subject[0] eq 'dfg05' )
			{
				$dfg_subject[0] = 'Multidisciplinary';
			}
			
			push (@dfg_subjects, $dfg_subject[0] );
		}
		else {
			push (@dfg_subjects, "" );
		}


		if( $eprint->exists_and_set( "oa_funding" ) )
		{
			# my $oa_funding = $eprint->get_value( 'oa_funding' );
			push (@oa_fundings, $oa_funding );
		}
		else {
			push (@oa_fundings, "" );
		}

});

##### Buffer ###########################################################################################
$results_buffer->map (sub
{
	my( $session, $dataset, $eprint ) = @_;
	
	my $type = $eprint->get_value( 'type' );
	my $oa_funding = $eprint->get_value( 'oa_funding' );

		#### Kontrollfeld
		my $kontrollfeld = '=WENN(ANZAHL2(B2:K2; N2:P2; R2)<14;"Pflichtfeld fehlt!";"ok")';
		push (@kontrollfelder, $kontrollfeld);

		#### eprintids
		if( $eprint->exists_and_set( "eprintid" ) )
		{		
			my $eprintid = $eprint->get_value( 'eprintid' );
			push (@eprintids, $eprintid);
		}		
		
		#### DOI
		if( $eprint->exists_and_set( "id_number" ) )
		{		
			my $doi = $eprint->get_value( 'id_number' );
			# $doi =~ s/^(.*)$/https:\/\/doi.org\/$1/;
			$doi =~ s/doi://;
			push (@dois, $doi);
		}
		else {
			push (@dois, "" );
		}

		#### Verlag
		if( $eprint->exists_and_set( "publisher" ) )
		{			
		my $publisher = $eprint->get_value( 'publisher' );
		
		push (@publishers, $publisher);
		}
		else {
			push (@publishers, "" );
		}		

		#### Publikationsform
		if( $eprint->exists_and_set( "type" ) )
		{			
			my $type = $eprint->get_value( 'type' );
			$type =~ s/article.*/journal article/;
			push (@types, $type);
		}
		else {
			push (@types, "" );
		}	

		#### CC-Lizenz
		my @documents = $eprint->get_all_documents();
		foreach my $doc ( @documents )
		{
			if( $doc->is_public )
			{
				if( $doc->exists_and_set( "license" ) )
				{
					my $license = $doc->get_value ('license');
					
					$license =~ s/cc_by_4/CC BY/;
					$license =~ s/cc_by_sa_4/CC BY-SA/;
					$license =~ s/cc_by_nd_4/CC BY-ND/;
					$license =~ s/cc_by_nc_4/CC BY-NC/;
					$license =~ s/cc_by_nc_sa_4/CC BY-NC-SA/;
					$license =~ s/cc_by_nc_nd_4/CC BY-NC-ND/;
					
					push (@licenses, $license);
				}
				else {
					push (@licenses, "" );
				}
			}
		}
		
		
		#### Waehrung
		if( $eprint->exists_and_set( "oa_apc_type" ) )
		{
			my $currency = $eprint->get_value( 'oa_apc_type' );
			push (@currencies, $currency);
		}
		else {
			push (@currencies, "" );
		}

		#### Rechnungsbetrag
		if( $eprint->exists_and_set( "oa_apc_value" ) )
		{		
			my $oa_apc_value = $eprint->get_value( 'oa_apc_value' );
			$oa_apc_value =~ s/\./,/;
			push (@oa_apc_values, $oa_apc_value);
		}
		else {
			push (@oa_apc_values, "" );
		}
		
		#### EURO Nettobetrag
		if( $eprint->exists_and_set( "oa_apc_net_value" ) )
		{		
			my $oa_apc_net_value = $eprint->get_value( 'oa_apc_net_value' );
			my $oa_apc_net_type = $eprint->get_value( 'oa_apc_net_type' );
			$oa_apc_net_value =~ s/\./,/;
			push (@oa_apc_net_values, $oa_apc_net_value);
			
			$net = $net + 1;
			$tax = $tax + 1;
			# =I20+(I20*J20)
			# Muss in der Endversion geändert werden auf
			# =H20+(H20*I20)
			# da die Spalte 'eprintid' entfällt
			push (@bruttos, "=I".$net."+(I".$net."*J".$tax.")");
		}
		else {
			push (@oa_apc_net_values, "" );
		}
		
		

		#### Steuersatz
		if( $eprint->exists_and_set( "oa_apc_tax_rate" ) )
		{		
			my $oa_apc_tax_rate = $eprint->get_value( 'oa_apc_tax_rate' );
			$oa_apc_tax_rate =~ s/\./,/;
			$oa_apc_tax_rate /= 100;
			push (@oa_apc_tax_rates, $oa_apc_tax_rate);
		}
		else {
			push (@oa_apc_net_values, "" );
		}

		### Foerdersumme
		if( $eprint->exists_and_set( "payment_amount" ) )
		{
			my $payment_amount =  $eprint->get_value( 'payment_amount' );
			push (@payment_amounts, $payment_amount );
		}
		else {
			push (@payment_amounts, "" );
		}

		# Transformationsvertraege
		if( $eprint->exists_and_set( "transformation" ) )
		{
		my $transformation_contract = $eprint->get_value( 'transformation' );
		$transformation_contract =~ s/no/kein/;
		$transformation_contract =~ s/springer_deal/Springer (DEAL)/;
		$transformation_contract =~ s/wiley_deal/Wiley (DEAL)/;
		push ( @transformation, $transformation_contract );
		}
		else {
			push (@transformation, "kein" );
		}
		
		# Foerderjahr
		if( $eprint->exists_and_set( "date" ) )
		{
			my $date = $eprint->get_value( 'date' );
			$date =~ s/^(\d{4}).*$/$1/;
			push ( @funding_year, $date );
		}		
		else {
			push (@funding_year, "" );
		}
	
		
		
		### DFG-Projektnummern
		if( $eprint->exists_and_set( "oa_funding_dfg_id" ) )
		{		
			my @dfg_id =  @{$eprint->get_value( 'oa_funding_dfg_id' )};
			push ( @dfg_number, $dfg_id[0] );
		}
		else {
			push (@dfg_number, "" );
		}
		

		### DFG-Wissenschaftsbereiche
		if( $eprint->exists_and_set( "dfg_subjects" ) )
		{
			my @dfg_subject =  @{$eprint->get_value( 'dfg_subjects' )};
		
			if ( $dfg_subject[0] eq 'dfg01' )
			{
				$dfg_subject[0] = 'Geistes- und Sozialwissenschaften';
			}
			if ( $dfg_subject[0] eq 'dfg02' )
			{
				$dfg_subject[0] = 'Lebenswissenschaften';
			}			
			if ( $dfg_subject[0] eq 'dfg03' )
			{
				$dfg_subject[0] = 'Naturwissenschaften';
			}	
			if ( $dfg_subject[0] eq 'dfg04' )
			{
				$dfg_subject[0] = 'Ingenieurwissenschaften';
			}	
			if ( $dfg_subject[0] eq 'dfg05' )
			{
				$dfg_subject[0] = 'Multidisciplinary';
			}
			
			push (@dfg_subjects, $dfg_subject[0] );
		}
		else {
			push (@dfg_subjects, "" );
		}


		if( $eprint->exists_and_set( "oa_funding" ) )
		{
			# my $oa_funding = $eprint->get_value( 'oa_funding' );
			push (@oa_fundings, $oa_funding );
		}
		else {
			push (@oa_fundings, "" );
		}




my $eprintid_ref = \@eprintids;
$worksheet->write_col( 1, 1, $eprintid_ref, $cells );

my $doi_ref = \@dois;
$worksheet->write_col( 1, 2, $doi_ref, $cells );

my $publisher_ref = \@publishers;
$worksheet->write_col( 1, 3, $publisher_ref, $cells );

my $type_ref = \@types;
$worksheet->write_col( 1, 4, $type_ref, $cells );

my $license_ref = \@licenses;
$worksheet->write_col( 1, 5, $license_ref, $cells );

my $oa_apc_type_ref = \@currencies;
$worksheet->write_col( 1, 6, $oa_apc_type_ref, $cells_center );

my $oa_apc_value_ref = \@oa_apc_values;
$worksheet->write_col( 1, 7, $oa_apc_value_ref, $cells );

my $oa_apc_net_value_ref = \@oa_apc_net_values;
$worksheet->write_col( 1, 8, $oa_apc_net_value_ref, $cells );

my $oa_apc_tax_rate_ref = \@oa_apc_tax_rates;
$worksheet->write_col( 1, 9, $oa_apc_tax_rate_ref, $cells_center );

my $brutto_ref = \@bruttos;
$worksheet->write_col( 1, 10, $brutto_ref, $cells );

my $payment_amount_ref = \@payment_amounts;
$worksheet->write_col( 1, 11, $payment_amount_ref, $cells );

my $transformation_ref = \@transformation;
$worksheet->write_col( 1, 14, $transformation_ref, $cells );

my $funding_year_ref = \@funding_year;
$worksheet->write_col( 1, 15, $funding_year_ref, $cells );

my $dfg_id_ref = \@dfg_number;
$worksheet->write_col( 1, 17, $dfg_id_ref, $cells );

my $dfg_subject_ref = \@dfg_subjects;
$worksheet->write_col( 1, 18, $dfg_subject_ref, $cells );

});

$workbook->close();
$session->terminate();
exit;

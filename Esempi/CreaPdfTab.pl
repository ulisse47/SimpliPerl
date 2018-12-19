#!/usr/bin/perl
 
 
#Installazione di PDF::API2 e PDF::Table
#sudo perl -MCPAN -e "install PDF::API2"
#sudo perl -MCPAN -e "install PDF::Table"
 
 
#Documentazione:
#http://search.cpan.org/~ssimms/PDF-API2-2.023/lib/PDF/API2.pm
#http://search.cpan.org/~omega/PDF-Table-0.9.7/lib/PDF/Table.pm
 
 
 
use PDF::API2;
use PDF::Table;
use Encode; #Per la codifica utf8 del testo
 
# Dati della tabella
 my $datiTabella =[
    ["Colonna0", "Colonna1","Colonna2"],
    ["cella0", "cella1", "cella2"],
    ["cella3", "cella4", "cella5"],
    ["cella6", "cella7", "cella8"],
    #... e così via...
 ];
 
 
    # Crea un nuovo pdf
    $pdf = PDF::API2->new();
 
 
    # Aggiunge una pagina bianca
    $page = $pdf->page();
 
    # Dimenzione pagina
    $page->mediabox('A4');
 
    # Font
    $font = $pdf->corefont('Helvetica-Bold');
 
    # Aggiungiamo un testo libero alla pagina
    $text = $page->text();
    $text->font($font, 20);
    $text->translate(2, 700);
    $text->text(decode("utf8",'Tabelle a volontà'));
 
 
    #Formattazione intestazione tabella
    my $intestazioneTabella = 
    {
        font       => $pdf->corefont("Helvetica", -encoding => "utf8"),
        font_size  => 18,
        font_color => '#004444',
        bg_color   => 'yellow', 
        repeat     => 1,    
        justify    => 'center'
    };
 
 
    #Se volessimo impostare le preferenze per una singola cella
    my $preferenzeCella = [];
    $preferenzeCella->[1][0] = {
        #Riga 2 cella 1
        background_color => '#008000',
        font_color       => '#FFFFFF',
        font_size  => 38,
    };
 
 
 
	# Creazione tabella
	my $pdftable = new PDF::Table;
	$pdftable->table(
		 # parametri obbligatori
		 $pdf,
		 $page,
		 $datiTabella,
		 header_props => $intestazioneTabella, 	#preferenze intestazione
		 cell_props => $preferenzeCella, 		#preferenze singola cella
		 x => 50,
		 w => 495,
		 start_y => 500,
		 start_h => 300,
		 # opzionali
		 next_y  => 750,
		 next_h  => 500,
		 padding => 5,
		 padding_right => 10,
		 border => 0 ,
		 background_color_odd  => "red", #sfondo righe dispari
		 background_color_even => "lightblue", #sfondo righe pari
	);
 
 # Salva il pdf PDF
    $pdf->saveas('documento-con-tabella.pdf');
 
    #Visualizza il pdf con il programma predefinito
    if ($^O=='MSWin32') { system('start documento-con-tabella.pdf')};
	if ($^O=='linux') { system('xdg-open documento-con-tabella.pdf')};

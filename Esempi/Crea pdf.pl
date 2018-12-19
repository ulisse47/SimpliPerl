#!/usr/bin/perl
 
# Installazione  PDF::API2
#sudo perl -MCPAN -e "install PDF::API2"
 
 
#Documentazione:
#http://search.cpan.org/~ssimms/PDF-API2-2.023/lib/PDF/API2.pm
 
 
 
use PDF::API2;
use Encode; #Per la codifica utf8 del testo
 
    # Crea un nuovo pdf
    $pdf = PDF::API2->new();
 
 
    # Aggiunge una pagina bianca
    $page = $pdf->page();
 
    # Dimenzione pagina
    $page->mediabox('A4');
 
    # Aggiunge un font del sistema
    $font = $pdf->corefont('Helvetica-Bold');
 
    # Aggiungi del testo 
    $text = $page->text();
    $text->font($font, 20);
    $text->translate(2, 700);
    # uso decode("utf8","bla bla...") per usare la codifica del testo utf8
    $text->text(decode("utf8",'Ciao mondo, saluti a volontà!')); 
 
    #Aggiunge un'immagine
    my $gfx=$page->gfx;
    $mypng = $pdf->image_png('logo.png');
    $gfx->image( $mypng, 30, 760 );
 
 
	# Salva il PDF
    $pdf->saveas('documento.pdf');
 
    #Visualizza il pdf con il programma predefinito
    if ($^O=='MSWin32') { system('start documento.pdf')};
	if ($^O=='linux') { system('xdg-open documento.pdf')};

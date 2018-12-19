use PDF::Reuse;
use strict;
    
    prDocDir('U:/Modelli');
    prFile('A4.pdf');
    prFont('Times-Roman');
    prCompress(1);     
	prForm('U:/Modelli/Pagamenti.pdf');                # Here the template is used
	prFont('Times-Roman');     # Just setting a font
	prFontSize(10);
    
	my ($fName, $lName) = ("Mau","Cav");
	
	prText(340,660, "Tit$fName $lName");
        prText(350,650, "Ragione Sociale");
        prText(350,630, "Email\@dominio");
        prText(70,525, "Tipo Documento");
        prText(240,525, "Ricevuta");
        prText(410,525, "Data pagamento");
        prText(240,485, "Ricevuta PayPal");
        prText(50,425, "Prodotto");
        prText(200,425, "Descrizione Prodotto");
        prText(420,425, "USD");
        prText(500,425, "Importo");
        prText(200,185, "Attivazione");

        prPage();

	prEnd();
    close INFILE;    

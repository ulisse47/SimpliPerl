for ($y=0;$y<1;$y++){

	$codelength =14; #//Lunghezza del codice (usare rand(min,max) per una lunghezza casuale)
	$salt = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	$code='';

	for($i=0;$i<=$codelength;$i++)	{

		$code.=substr($salt,rand(length($salt)),1);

		}

	print   '0,' .
		"SMY18$code". ',' .
		"NULL" . ',' . 
		'"00",'  .
		'"2018-12-31 00:00:00",' .
		'' . ','.
		',' .
		',' .
		'"0000-00-00 00:00:00",'.
		'0,' .
		'""' .
		"\n";
	}

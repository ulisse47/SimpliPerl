for ($y=0;$y<50;$y++){
	$codelength =14; #//Lunghezza del codice (usare rand(min,max) per una lunghezza casuale)
	$salt = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	$code='';
	for($i=0;$i<=$codelength;$i++)	{
		$code.=substr($salt,rand(length($salt)),1);
		}

	print   '0,' .
		"PON99$code". ',' .
		NULL . ',' . 
		'"00",'  .
		'"2018-10-01 00:00:00",' .
		',' .
		',' .
		',' .
		'"0000-00-00 00:00:00",'.
		'0,' .
		'"SerFL"' .
		"\n";
	}

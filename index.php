<?php
ini_set('display_errors', 1);
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

if (!is_file($autoloadFile = __DIR__ . '/vendor/autoload.php')) {
	throw new \LogicException('Could not find autoload.php in vendor/');
}
require $autoloadFile;
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
use Digitick\Sepa\DomBuilder\CustomerDirectDebitTransferDomBuilder;
use Digitick\Sepa\Exception\InvalidTransferFileConfiguration;
use Digitick\Sepa\GroupHeader;
use Digitick\Sepa\PaymentInformation;
use Digitick\Sepa\TransferFile\CustomerDirectDebitTransferFile;
use Digitick\Sepa\TransferInformation\CustomerDirectDebitTransferInformation;

function isValidIBAN ($iban) {

	if (preg_match('/^[A-Z]{2}[0-9]{2}[A-Z0-9]{1,30}$/', $iban)) {

		$iban = strtolower($iban);
		$Countries = array(
			'al'=>28,'ad'=>24,'at'=>20,'az'=>28,'bh'=>22,'be'=>16,'ba'=>20,'br'=>29,'bg'=>22,'cr'=>21,'hr'=>21,'cy'=>28,'cz'=>24,
			'dk'=>18,'do'=>28,'ee'=>20,'fo'=>18,'fi'=>18,'fr'=>27,'ge'=>22,'de'=>22,'gi'=>23,'gr'=>27,'gl'=>18,'gt'=>28,'hu'=>28,
			'is'=>26,'ie'=>22,'il'=>23,'it'=>27,'jo'=>30,'kz'=>20,'kw'=>30,'lv'=>21,'lb'=>28,'li'=>21,'lt'=>20,'lu'=>20,'mk'=>19,
			'mt'=>31,'mr'=>27,'mu'=>30,'mc'=>27,'md'=>24,'me'=>22,'nl'=>18,'no'=>15,'pk'=>24,'ps'=>29,'pl'=>28,'pt'=>25,'qa'=>29,
			'ro'=>24,'sm'=>27,'sa'=>24,'rs'=>22,'sk'=>24,'si'=>19,'es'=>24,'se'=>24,'ch'=>21,'tn'=>24,'tr'=>26,'ae'=>23,'gb'=>22,'vg'=>24
		);
		$Chars = array(
			'a'=>10,'b'=>11,'c'=>12,'d'=>13,'e'=>14,'f'=>15,'g'=>16,'h'=>17,'i'=>18,'j'=>19,'k'=>20,'l'=>21,'m'=>22,
			'n'=>23,'o'=>24,'p'=>25,'q'=>26,'r'=>27,'s'=>28,'t'=>29,'u'=>30,'v'=>31,'w'=>32,'x'=>33,'y'=>34,'z'=>35
		);

		if (strlen($iban) != $Countries[ substr($iban,0,2) ]) { return false; }

		$MovedChar = substr($iban, 4) . substr($iban,0,4);
		$MovedCharArray = str_split($MovedChar);
		$NewString = "";

		foreach ($MovedCharArray as $k => $v) {

			if ( !is_numeric($MovedCharArray[$k]) ) {
				$MovedCharArray[$k] = $Chars[$MovedCharArray[$k]];
			}
			$NewString .= $MovedCharArray[$k];
		}
		if (function_exists("bcmod")) { return bcmod($NewString, '97') == 1; }

		// http://au2.php.net/manual/en/function.bcmod.php#38474
		$x = $NewString; $y = "97";
		$take = 5; $mod = "";

		do {
			$a = (int)$mod . substr($x, 0, $take);
			$x = substr($x, $take);
			$mod = $a % $y;
		}
		while (strlen($x));

		return (int)$mod == 1;
	}
	else{return false;}
}

function findPmntKey($sheet){
	$highestRow = $sheet->getHighestRow();
	$highestColumn = $sheet->getHighestColumn();
	$columncount = PHPExcel_Cell::columnIndexFromString($highestColumn);
	$key = null;
	for ($row = 0; $row <= $highestRow - 2; $row++) {
		for ($column = 0; $column <= $columncount - 1; $column++) {

			if(strpos($sheet->getCellByColumnAndRow($column,$row),'PmtInfId')!==false)
			{
				$key = $row;
				break;
			}
		}
		if($key!==null)
		{
			break;
		}
	}
	return $key;
}
function findPivot($sheet,$values){
	$highestRow = $sheet->getHighestRow();
	$highestColumn = $sheet->getHighestColumn();
	$columncount = PHPExcel_Cell::columnIndexFromString($highestColumn);
	$colStart = $values['col_start'];
	$colEnd = $values['col_end'];
	$key = $values['key'];
	for ($row = 0; $row <= $highestRow - 2; $row++) {
		for ($column = $colStart; $column <= $colEnd - 1; $column++) {

			if(strpos($sheet->getCellByColumnAndRow($column,$row)->getValue(),$key)!==false)
			{
				return $row;
			}
		}
	}
}
function getPaymentsFromExcel($sheet)
{
	$pmnts = array();
	$highestRow = $sheet->getHighestRow();
	$highestColumn = $sheet->getHighestColumn();
	$columncount = PHPExcel_Cell::columnIndexFromString($highestColumn);
	$index = findPmntKey($sheet);
	$pivot = $index + 1;
	for ($row = $pivot; $row <= $highestRow - 2; $row++) {
		for ($column = 0; $column <= $columncount - 1; $column++) {
			$val = $sheet->getCellByColumnAndRow($column,$row)->getCalculatedValue();
			if($val===null){
				break;
			}
			$pmnts[$row][$sheet->getCellByColumnAndRow($column,$index)->getValue()] = $val;
		}
	}
	$pmnts = array_values($pmnts);
	return $pmnts;
}

function getHeaderFromExcel($sheet){
	$header = array();
	for ($row = 2; $row < 3; $row++) {
		for ($column = 31; $column <= 52 - 1; $column++) {
			$header[$row][$sheet->getCellByColumnAndRow($column,2)->getValue()] = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();
		}
	}
	$header = array_values($header);
	return $header[0];
}
function getHeaderCtrSum($sheet){
	$headerCtrSum = array();
	for ($row = 2; $row < 3; $row++) {
		for ($column = 33; $column <= 35 - 1; $column++) {
			$headerCtrSum[$row][$sheet->getCellByColumnAndRow($column,2)->getValue()] = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();
		}
	}
	$headerCtrSum = array_values($headerCtrSum);
	return $headerCtrSum[0];
}
function getRcurCtrSum($sheet){
	$header = array();
	for ($row = 2; $row < 3; $row++) {
		for ($column = 40; $column <= 42 - 1; $column++) {
			$header[$row][$sheet->getCellByColumnAndRow($column,2)->getValue()] = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();
		}
	}
	$header = array_values($header);
	return $header[0];
}

function getFrstCtrSum($sheet){
	$header = array();
	$pivot = findPivot($sheet,array('col_start'=>55,'col_end'=>57,'key'=>'NbOfTxs'));
	for ($row = $pivot; $row < 9; $row++) {
		for ($column = 55; $column <= 57 - 1; $column++) {			
				$header[$row][$sheet->getCellByColumnAndRow($column,$pivot)->getValue()] = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();	
						
		}
	}
	$header = array_values($header);
	return $header[0];
}


function getRcursFromExcel($sheet){
	$rcurs = array();
	$pivot = findPivot($sheet,array('col_start'=>2,'col_end'=>15,'key'=>'EndToEndId'));
	$highestRow = $sheet->getHighestRow();
	for ($row = $pivot; $row < $highestRow - 2; $row++) {
		for ($column = 2; $column <= 15 - 1; $column++) {
				$cell = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();
				if($cell !== NULL && $cell !== '')
				{
					$rcurs[$row][$sheet->getCellByColumnAndRow($column,$pivot)->getValue()] = $cell;
				}
		}
	}
	$rcurs = array_values($rcurs);
	return $rcurs;
}

function getFrstFromExcel($sheet){
	$frst = array();
	$pivot = findPivot($sheet,array('col_start'=>17,'col_end'=>30,'key'=>'EndToEndId'));
	$highestRow = $sheet->getHighestRow();
	for ($row = $pivot; $row < $highestRow - 2; $row++) {
		for ($column = 17; $column <= 30 - 1; $column++) {
			$cell = $sheet->getCellByColumnAndRow($column,$row +1)->getCalculatedValue();
			if($cell !== NULL && $cell !== '')
			{
				$frst[$row][$sheet->getCellByColumnAndRow($column,$pivot)->getValue()] = $cell;
			}
		}
	}
	$frst = array_values($frst);
	return $frst;
}

function getPaymentHeader($header,$pmnt)
{	
	$pmntheader = array();			
	$pmntheader['MsgId'] = $header['MsgId'];
	$pmntheader['CdtrAcct-IBAN'] = $pmnt['CdtrAcct-IBAN'];
	$pmntheader['CdtrAgt-BIC'] = $pmnt['CdtrAgt-BIC'];
	$pmntheader['Nm'] = $header['Nm'];
	$pmntheader['SeqTp'] = $pmnt['SeqTp'];
	$pmntheader['ReqdColltnDt'] = $pmnt['ReqdColltnDt'];
	$pmntheader['Id'] = $pmnt['Id'];
	//$pmntheader['CtgyPurp'] = $pmnt['CtgyPurp'];
	$pmntheader['AdrLine'] = $pmnt['AdrLine'];			
	return $pmntheader;
}

function formatPrice($price){
    $CtrlSum = 0;
    if(strpos($price, ',')!==FALSE)
    {
        $CtrlSum = str_replace(',','.',$price);
    }
    elseif(strpos($CtrlSum, '.')!==FALSE)
    {
        $CtrlSum = $price;
    }
    else{
        $CtrlSum = $price.'.00';
    }
    $CtrlSum = number_format($CtrlSum, 2, '.', '');
    return $CtrlSum;
}

function generatePayment($header,$pmnt){
	$payment = new PaymentInformation($header['MsgId'].'0', $pmnt['CdtrAcct-IBAN'], $pmnt['CdtrAgt-BIC'], $header['Nm']);
	if($pmnt['SeqTp']==='RCUR')
	{
		$payment->setSequenceType(PaymentInformation::S_RECURRING);
	}
	else{
		$payment->setSequenceType(PaymentInformation::S_FIRST);
	}
	$payment->setDueDate(new \DateTime($pmnt['ReqdColltnDt']));
	$payment->setCreditorId($pmnt['Id']);
	$payment->setCountry(substr($pmnt['CdtrAcct-IBAN'],0,2));
	$payment->setAddressLine(array($pmnt['AdrLine']));
	//$payment->setCtgyPurposeCode($pmnt['CtgyPurp']);
	$transfer = new CustomerDirectDebitTransferInformation(formatPrice($pmnt['CtrlSum']), $pmnt['DbtrAcct-IBAN'], $pmnt['Dbtr-Nm']);
	$transfer->setBic($pmnt['DbtrAgt-BIC']);
	$transfer->setEndToEndIdentification($pmnt['EndToEndId']);
	$transfer->setMandateSignDate(new \DateTime($pmnt['DtOfSgntr']));
	$transfer->setMandateId($pmnt['MndtId']);
	$transfer->setRemittanceInformation($pmnt['Ustrd']);
	$transfer->setCountry(substr($pmnt['CdtrAcct-IBAN'], 0, 2));
	$transfer->setAddressLine(array($pmnt['Dbtr-AdrLine']));
	//$transfer->setPurposeCode($pmnt['Purp']);
	$payment->addTransfer($transfer);
	return $payment;
}
function generateTransfer($pmnt){	
	$transfer = new CustomerDirectDebitTransferInformation(formatPrice($pmnt['CtrlSum']), $pmnt['DbtrAcct-IBAN'], $pmnt['Dbtr-Nm']);
	$transfer->setBic($pmnt['DbtrAgt-BIC']);
	$transfer->setEndToEndIdentification($pmnt['EndToEndId']);
	$transfer->setMandateSignDate(new \DateTime($pmnt['DtOfSgntr']));
	$transfer->setMandateId($pmnt['MndtId']);
	$transfer->setRemittanceInformation($pmnt['Ustrd']);
	$transfer->setCountry(substr($pmnt['CdtrAcct-IBAN'], 0, 2));
	$transfer->setAddressLine(array($pmnt['Dbtr-AdrLine']));
	//$transfer->setPurposeCode($pmnt['Purp']);	
	return $transfer;
}

function getRcurPayments($pmnts){
	$transfers = array();
	foreach ($pmnts as $pmnt) {

				if(isset($pmnt['CtrlSum']))
				{

					$transfer = new CustomerDirectDebitTransferInformation(formatPrice($pmnt['CtrlSum']), $pmnt['DbtrAcct-IBAN'], $pmnt['Dbtr-Nm']);
				$transfer->setBic($pmnt['DbtrAgt-BIC']);
				$transfer->setEndToEndIdentification($pmnt['EndToEndId']);
				$transfer->setMandateSignDate(new \DateTime($pmnt['DtOfSgntr']));
				$transfer->setMandateId($pmnt['MndtId']);
				$transfer->setRemittanceInformation($pmnt['Ustrd']);
				$transfer->setCountry(substr($pmnt['DbtrAcct-IBAN'], 0, 2));
				$transfer->setAddressLine(array($pmnt['Dbtr-AdrLine']));
				//$transfer->setPurposeCode($pmnt['Purp']);	
				$transfers[]=  $transfer;
				}
				
	}
	return $transfers;
}

function getFrstPayments($pmnts){
	$transfers = array();
	foreach ($pmnts as $pmnt) {
			

				if(isset($pmnt['CtrlSum'])){

					$transfer = new CustomerDirectDebitTransferInformation(formatPrice($pmnt['CtrlSum']), $pmnt['DbtrAcct-IBAN'], $pmnt['Dbtr-Nm']);
				$transfer->setBic($pmnt['DbtrAgt-BIC']);
				$transfer->setEndToEndIdentification($pmnt['EndToEndId']);
				$transfer->setMandateSignDate(new \DateTime($pmnt['DtOfSgntr']));
				$transfer->setMandateId($pmnt['MndtId']);
				$transfer->setRemittanceInformation($pmnt['Ustrd']);
				$transfer->setCountry(substr($pmnt['DbtrAcct-IBAN'], 0, 2));
				$transfer->setAddressLine(array($pmnt['Dbtr-AdrLine']));
				//$transfer->setPurposeCode($pmnt['Purp']);	
				$transfers[]=  $transfer;
				}				

	}
	return $transfers;
}

function generateSepaXml($inputFileName){
	try {
		$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objReader->setPassword('01');
		$objPHPExcel = $objReader->load($inputFileName);
	} catch(Exception $e) {
		die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
	}
	$sheet = $objPHPExcel->getSheet(0);
	$header = getHeaderFromExcel($sheet);
	$headerCtrSum = getHeaderCtrSum($sheet);
	$rcurCtrSum = getRcurCtrSum($sheet);
	$frstCtrSum = getFrstCtrSum($sheet);
	$rs = getRcursFromExcel($sheet);	
	$fs = getFrstFromExcel($sheet);
	//$pmnts = getPaymentsFromExcel($sheet);
	//$pmnth = getPaymentHeader($header,$pmnts[0]);
	$rcurpayments = getRcurPayments($rs);
	$frstpayments = getFrstPayments($fs);
	$groupHeader = new GroupHeader($header['MsgId'], $header['Nm']);
	$groupHeader->setCreationDateTime($header['CreDtTm']);
	$groupHeader->setNumberOfTransactions($headerCtrSum['NbOfTxs']);
	$groupHeader->setControlSumCents(formatPrice($headerCtrSum['CtrlSum']));
	$groupHeader->setInitiatingPartyId('ES68000B70440664');
	$sepaFile = new CustomerDirectDebitTransferFile($groupHeader);
	$rpayment = null;
	$fpayment = null;
	foreach ($rcurpayments as $rp) {
				if(!$rpayment)
				{
					$rpayment = new PaymentInformation($header['MsgId'], $header['CdtrAcct-IBAN'], $header['CdtrAgt-BIC'], $header['Nm']);
					$rpayment->setDueDate(new \DateTime($header['ReqdColltnDt']));
					$rpayment->setCreditorId($header['Id']);
					$rpayment->setCountry(substr($header['CdtrAcct-IBAN'],0,2));
					$rpayment->setAddressLine(array($header['AdrLine']));
					//$rpayment->setCtgyPurposeCode($header['CtgyPurp']);
					$rpayment->setControlSumCents(formatPrice($rcurCtrSum['CtrlSum']));
					$rpayment->setNumberOfTransactions($rcurCtrSum['NbOfTxs']);
					$rpayment->setSequenceType(PaymentInformation::S_RECURRING);
				}
				$rpayment->addTransfer($rp);
			}
			if($rpayment)
            {
                $sepaFile->addPaymentInformation($rpayment);
            }
	foreach ($frstpayments as $fp) {
				if(!$fpayment)
				{
					$fpayment = new PaymentInformation($header['MsgId'].'0', $header['CdtrAcct-IBAN'], $header['CdtrAgt-BIC'], $header['Nm']);
					$fpayment->setDueDate(new \DateTime($header['ReqdColltnDt']));
					$fpayment->setCreditorId($header['Id']);
					$fpayment->setCountry(substr($header['CdtrAcct-IBAN'],0,2));
					$fpayment->setAddressLine(array($header['AdrLine']));
					//$fpayment->setCtgyPurposeCode($header['CtgyPurp']);
					$fpayment->setControlSumCents(formatPrice($frstCtrSum['CtrlSum']));
					$fpayment->setNumberOfTransactions($frstCtrSum['NbOfTxs']);
					$fpayment->setSequenceType(PaymentInformation::S_FIRST);
				}
				$fpayment->addTransfer($fp);
			}
			if($fpayment)
            {
                $sepaFile->addPaymentInformation($fpayment);
            }
	 
	/*foreach ($pmnts as $pmnt)
	{		
		if(!in_array($pmnt['SeqTp'], $seqs))
		{
			$payment = generatePayment($header,$pmnt);
			$seqs[] = $pmnt['SeqTp'];
		}
		else{
			$payment->addTransfer(generateTransfer($pmnt));
		}
		
		
	}*/
	$domBuilder = new CustomerDirectDebitTransferDomBuilder('pain.008.001.02');
	$sepaFile->accept($domBuilder);
	$xml = $domBuilder->asXml();
	$dom = new \DOMDocument('1.0', 'UTF-8');
	$dom->loadXML($xml);
	$creationDate = date('d-m-Y H:i:s');
	$dom->save("Recibos.xml");
	if (file_exists("Recibos.xml")) {
	    header('Content-Description: File Transfer');
	    header('Content-Type: application/octet-stream');
	    header('Content-Disposition: attachment; filename="'.basename("Recibos.xml").'"');
	    header('Expires: 0');
	    header('Cache-Control: must-revalidate');
	    header('Pragma: public');
	    header('Content-Length: ' . filesize("Recibos.xml"));
	    readfile("Recibos.xml");
	    exit;
	}
}
generateSepaXml('Recibos.xls');
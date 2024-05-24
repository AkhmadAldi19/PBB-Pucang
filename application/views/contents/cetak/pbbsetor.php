<?php

$pdf = new TCPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);
// set document information
$pdf->SetCreator(PDF_CREATOR);
$pdf->SetAuthor('Rii');
$pdf->SetTitle('Laporan PBB Setor');
$pdf->SetSubject('Laporan PBB Setor');
$pdf->SetKeywords('PBB Setor');

// set default header data
//$pdf->SetHeaderData(PDF_HEADER_LOGO, PDF_HEADER_LOGO_WIDTH, PDF_HEADER_TITLE.' ', PDF_HEADER_STRING);
$pdf->setFooterData();

// set header and footer fonts
$pdf->setHeaderFont(Array(PDF_FONT_NAME_MAIN, '', PDF_FONT_SIZE_MAIN));
$pdf->setFooterFont(Array(PDF_FONT_NAME_DATA, '', PDF_FONT_SIZE_DATA));

// set default monospaced font
$pdf->SetDefaultMonospacedFont(PDF_FONT_MONOSPACED);

// set margins
$pdf->SetMargins(PDF_MARGIN_LEFT, PDF_MARGIN_TOP, PDF_MARGIN_RIGHT);
$pdf->SetHeaderMargin(PDF_MARGIN_HEADER);
$pdf->SetFooterMargin(PDF_MARGIN_FOOTER);

// set auto page breaks
$pdf->SetAutoPageBreak(TRUE, PDF_MARGIN_BOTTOM);

// set image scale factor
$pdf->setImageScale(PDF_IMAGE_SCALE_RATIO);

// set some language-dependent strings (optional)
if (@file_exists(dirname(__FILE__).'/lang/eng.php')) {
    require_once(dirname(__FILE__).'/lang/eng.php');
    $pdf->setLanguageArray($l);
}

// ---------------------------------------------------------

// set default font subsetting mode
$pdf->setFontSubsetting(true);

$pdf->AddPage();
ob_start();

$tblHeader =  <<<EOD
<style>
    .pbbsetor:border:1px;
</style>
<hr>
<h1 style="text-align:center;">LAPORAN PBB SETOR DESA PUCANGSARI</h1>

<table cellspacing="0" cellpadding="1" border="1" width="100%"> 
<thead >
    <tr style="font-weight:bold; text-align:center;">
        <th width="5%">No</th>
        <th width="30%">NOP</th>
        <th width="35%">NAMA WP</th>
        <th width="30%">PAJAK (Rp.</th>
    </tr>
</thead>
<tbody>
EOD;

$tbl ='';
$total = 0;

if($datas->num_rows() >0){
    $no = 0;
    foreach($datas->result() as $row){                                     $no++;         
    $tbl .= <<<EOD
    <tr>
        <td width="5%" style="text-align:center;">{$no}</td>
        <td width="30%"> {$row->nop} </td> 
        <td width="35%">{$row->nama} </td>  
        <td width="30%"> {$rupiah($row->total_pajak) }</td>  
    </tr>
    EOD;
    $total =+ $row->total_pajak;
    }
}
$tblFooter = <<<EOD
</tbody>
<tfoot>
    <tr style="font-weight:bold;">
        <td colspan="3">Total</td>
        <td>{$rupiah($total)}</td>
    </tr>
</tfoot>
</table>
<h4 >NB : Setoran dari tanggal {$tgl_awal} sampai dengan {$tgl_akhir}</h4>

<br>
<br>
        <p align="right">Pucangsari, {$tgl_akhir}</p>
        <p align="left">Mengesahkan</p>
        <br>
        <p align="left">____________________</p>
        <p align="left">Kepala Desa Pucangsari</p>
EOD;


ob_end_clean();
$pdf->writeHTML($tblHeader . $tbl . $tblFooter, true, false, true, false, '');
$pdf->Output('output.pdf', 'I');                    
?>
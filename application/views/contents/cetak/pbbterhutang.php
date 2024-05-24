<?php
$pdf = new TCPDF('L', PDF_UNIT, 'F4', true, 'UTF-8', false);
// set document information
$pdf->SetCreator(PDF_CREATOR);
$pdf->SetAuthor('Kadiyanto');
$pdf->SetTitle('Laporan PBB Terhutang');
$pdf->SetSubject('Laporan PBB Terhutang');
$pdf->SetKeywords('PBB Terhutang');

// set default header data
// $pdf->SetHeaderData(PDF_HEADER_LOGO, PDF_HEADER_LOGO_WIDTH, PDF_HEADER_TITLE . ' ', PDF_HEADER_STRING);

// set header and footer fonts
$pdf->setHeaderFont(array(PDF_FONT_NAME_MAIN, '', PDF_FONT_SIZE_MAIN));
$pdf->setFooterFont(array(PDF_FONT_NAME_DATA, '', PDF_FONT_SIZE_DATA));

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
if (@file_exists(dirname(__FILE__) . '/lang/eng.php')) {
    require_once(dirname(__FILE__) . '/lang/eng.php');
    $pdf->setLanguageArray($l);
}

// ---------------------------------------------------------

// set default font subsetting mode
$pdf->setFontSubsetting(true);

$pdf->AddPage();
ob_start();

$tblHeader =  <<<EOD
<style>
    .pembayaran:border:1px;
</style>
<hr>
<h1 style="text-align:center;">LAPORAN BELUM BAYAR PBB DESA PUCANGSARI</h1>
<h1 style="text-align:center;"></h1>
<table width="60%">
<table cellspacing="0" cellpadding="1" border="1" width="100%"> 
<thead >
    <tr style="font-weight:bold; text-align:center;">
        <th width="5%">No</th>
        <th width="15%">Nop</th>
        <th width="19%">Nama WP</th>
        <th width="20%">Alamat WP</th>
        <th width="10%">Alamat Objek</th>
        <th width="8%">Luas Bumi</th>
        <th width="8%">Luas Bangunan</th>

        <th width="15%">Pajak</th>
    </tr>
</thead>
<tbody>
EOD;

$tbl = '';
$total = 0;

if ($datas->num_rows() > 0) {
    $no = 0;
    foreach ($datas->result() as $row) {
        $no++;
        $tbl .= <<<EOD
    <tr >
        <td width="5%" style=" text-align:center;">{$no}</td>
        <td width="15%"> {$row->nop} </td> 
        <td width="19%">{$row->nama} </td>  
        <td width="20%">{$row->alamat_wp} </td>  
        <td width="10%">{$row->alamat_op} </td>  
        <td width="8%">{$row->bumi} </td>  
        <td width="8%">{$row->bangunan} </td>  
        <td width="15%"> {$rupiah($row->pajak)}</td>  
    </tr>
    EOD;
        $total = +$row->pajak;
    }
}
$tblFooter = <<<EOD
</tbody>
<tfoot>
    <tr style="font-weight:bold; ">
        <td colspan="3">Total</td>
        <td>{$rupiah($total)}</td>
    </tr>
</tfoot>
</table>
EOD;

ob_end_clean();
$pdf->writeHTML($tblHeader . $tbl . $tblFooter, true, false, true, false, '');
$pdf->Output('output.pdf', 'I');

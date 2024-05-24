<?php


$pdf = new TCPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);
// set document information
$pdf->SetCreator(PDF_CREATOR);
$pdf->SetAuthor('Siap Ayem');
$pdf->SetTitle('TANDA TERIMA PBB');
$pdf->SetSubject('TANDA TERIMA PBB');
$pdf->SetKeywords('PBB tanda terima');

// $pdf->SetHeaderData('/../images/pasuruan.png', PDF_HEADER_LOGO_WIDTH, 'Desa Pucangsari', 'Kecamatan Purwosari, Pasuruan');
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


// set some language-dependent strings (optional)
if (@file_exists(dirname(__FILE__) . '/lang/eng.php')) {
    require_once(dirname(__FILE__) . '/lang/eng.php');
    $pdf->setLanguageArray($l);
}

// ---------------------------------------------------------

// set default font subsetting mode
$pdf->setFontSubsetting(true);

$pdf->AddPage();


if ($data->num_rows() > 0) {
    $d = $data->row();

    $html =  <<<EOD
<style>
    .pembayaran:border:1px;
</style>
<hr>

<h1 style="text-align:center; margin-bottom: 20px;">TANDA TERIMA PBB PUCANGSARI</h1>

<br>
<table width="60%" style="margin-top: 20px;">
<br>
<tbody style="margin-top: 20px;">
    <tr>
        <td width="40%">NOP</td>
        <td width="5%">:</td>
        <td width="55%">{$d->nop}</td>
    </tr>
    <tr>
        <td>Nama Wajib Pajak</td>
        <td>:</td>
        <td>{$d->nama}</td>
    </tr>
    <tr>
        <td>Alamat Wajib Pajak</td>
        <td>:</td>
        <td>{$d->alamat_wp}</td>
    </tr>
    <tr>
        <td>Alamat Objek Pajak</td>
        <td>:</td>
        <td>{$d->alamat_op}</td>
    </tr>
    <tr>
        <td>Luas Bumi</td>
        <td>:</td>
        <td>{$d->bumi} m2</td>
    </tr>
    <tr>
        <td>Luas Bangunan</td>
        <td>:</td>
        <td>{$d->bangunan} m2</td>
    </tr>
    <tr>
        <td>Status</td>
        <td>:</td>
        <td>{$d->ket}</td>
    </tr>
</tbody>
</table>
EOD;
    $html2 = '';
    if ($d->ket == "Lunas") {
        $html2 .= <<<EOD
    <div class="pembayaran">
        <h3>Pembayaran Pajak</h3>
        <p>Pajak : {$rupiah($d->pajak)}</p>
        <p>Denda : {$rupiah(!empty($d->denda) ?$d->denda : 0)}</p>
        <p>Total Pembayaran : {$rupiah(!empty($d->total_pajak) ?$d->total_pajak :$d->pajak)}</p>
        <p>Tanggal Pembayaran : {$tanggal($d->tanggal)}</p>
        <br>
        <p align="left">Petugas</p>
        <br>
        <p align="left">____________________</p>
        <p align="left">{$d->nama_petugas}</p>
        
    </div>
 EOD;
    }
}
$pdf->writeHTML($html . $html2, true, false, true, false, '');
$pdf->Output('output.pdf', 'I');



<!-- ~~~~~~~~~~~~ HEADER ~~~~~~~~~~~ -->
<?=$this->load->view('layouts/header','',TRUE)?>

    <div class="page-wrapper">
    
    <!-- ~~~~~~~~~~~~ SIDEBAR ~~~~~~~~~~~ -->
    <?=$this->load->view('layouts/sidebar','',TRUE)?>


        <!-- PAGE CONTAINER-->
        <div class="page-container2">
            <!-- ~~~~~~~~~~~~ TOPNAV CONTENT ~~~~~~~~~~~ -->
            <?=$this->load->view('layouts/topnavcontent','',TRUE)?>

            <!-- ~~~~~~~~~~~~ TOPNAV CONTENT ~~~~~~~~~~~ -->
            <!-- MAIN CONTENT-->
            <div class="main-content">
                <div class="section__content section__content--p30">
            
                    <div class="container-fluid">
                       <!-- <?=base_url('Welcome/index')?>
                    <input type="file" name="excel" required />
                    <p class="help">* Gunakan file dengan extensi .xlsx</p>
                     <button type="submit" name="submit" value="upload">Upload File Excell </button>
                <?= form_close(); ?> -->
                <form action="<?=base_url('Welcome/index')?>" enctype="multipart/form-data" method="post">
                    <input type="file" name="upload_excel" required />
                    <button class="m-t-10 btn-success" type="submit" name="submit" value="upload">Upload File Excell </button>
                    <!-- <input type="submit" name="submit" value="Submit"> -->
                    <!-- <?php if($this->session->flashdata('success'))  { ?>
                        <p><?=$this->session->flashdata('success')?></p>
                    <?php  } ?>
                    <?php if($this->session->flashdata('error'))  { ?>
                        <p><?=$this->session->flashdata('error')?></p>
                    <?php  } ?> -->
	            </form>
            </div>
            <hr>

<div class="row">
    <div class="col-md-12">
        <!-- DATA TABLE -->
        <h3 class="title-5 m-b-35">Syarat dan Ketentuan Uplod Data DHKP Excell</h3>
        
                
                <p>1. Gunakan Ms.Office Excell untuk input data</p>
                <p>2. Isi pada baris pertama dengan angka nol</p>
                <p>3. Isi kolom A dengan Nomor Urut pada DHKP</p>
                <p>4. Isi kolom B dengan Nomor Objek Pajak (NOP)</p>
                <p>5. Isi kolom C dengan Nama Wajib Pajak </p>
                <p>6. Isi kolom D dengan Luas Bumi</p>
                <p>7. Isi kolom E dengan Luas Bangunan</p>
                <p>8. Isi kolom F dengan Nominal Pajak PBB</p>
                <p>9. Isi kolom G dengan Alamat Objek Pajak</p>
                <p>10. Isi kolom H dengan Alamat Wajib Pajak</p>
                <p>11. Isi kolom I dengan keterangan = Belum Bayar </p>
                <p>Lihat Contoh Pengisian dibawah ini </p>
                </div>
                 <img src="<?=site_url()?>uploads/template/importdhkp.png" alt="SIAP AYEM">  

                </div>
                                
                               
                            

    <style>
        body {
            padding: 20px 50px;
        }

        input[type="file"] {
            display: block;
            border: 1px solid #b5b5b5;
            border-radius: 2px;
            padding: 5px;
            font-size: 16px;
            color: #777;
        }

        button[type="submit"] {
            background: #06c;
            border-color: #06c;
            outline: none;
            color: #fff;
            font-size: 15px;
            cursor: pointer;
            border-radius: 3px;
            padding: 5px 10px;
        }

        .help {
            font-size: 12px;
            color: #b5b5b5;
            margin-top: 2px;
            font-style: italic;
        }
    </style>
                        
                    </div>
                </div>
            </div>

        </div>

    </div>

<!-- ~~~~~~~~~~~~ SCRIPT ~~~~~~~~~~~ -->
<?=$this->load->view('layouts/script','',TRUE)?>

<?php 
if(!empty($script)){
   echo $this->load->view($script,'',TRUE);
}
?>

<!-- ~~~~~~~~~~~~ FOOTER ~~~~~~~~~~~ -->
<?=$this->load->view('layouts/footer','',TRUE)?>


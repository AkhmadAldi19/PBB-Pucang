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
                <div class="section__content section__content--p30 ">
                    <div class="container-fluid">
                        <?=$this->load->view(!empty($content)?$content:'contents','',TRUE)?> 

                        
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

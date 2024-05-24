<!-- halaman Edit data sppt -->
<div class="row">

<div class="col-lg-12">
    <div class="card">
        <div class="card-header">
            <strong><i class="fa fa-edit"></i> Edit Data SPPT</strong>
        </div>
        <?php 
        
        if(isset($data)){
            $d['d'] = $data->row();
        ?>
        <form action="<?=base_url()?>datasppt/editSave/<?=$d['d']->id?>" method="post" class="form-horizontal">
            <!-- form Edit data sppt -->
            <div class="card-body card-block">
                    <?=$this->load->view('contents/datasppt/form', $d,TRUE)?>
            </div>
            <div class="card-footer">
                <!-- Tombol Edit data sppt -->
                <button type="submit" class="btn btn-primary btn-sm">SUBMIT
                </button>
            </div>
        </form>
        <?php } ?>
    </div>
</div>
</div>
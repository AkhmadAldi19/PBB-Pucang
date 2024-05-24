
<div class="row">

<div class="col-lg-12">
    <div class="card">
        <div class="card-header">
            <strong><i class="fa fa-edit"></i> Edit Petugas PBB</strong>
        </div>
        <?php 
        
        if(isset($data)){
            $d['d'] = $data->row();
        ?>
        <form action="<?=base_url()?>petugas/editSave/<?=$d['d']->id?>" method="post" class="form-horizontal">
            <div class="card-body card-block">
                    <?=$this->load->view('contents/petugas/form', $d,TRUE)?>
            </div>
            <div class="card-footer">
                <button type="submit" class="btn btn-primary btn-sm">PERBARUI
                </button>
            </div>
        </form>
        <?php } ?>
    </div>
</div>
</div>

<div class="row">

<div class="col-lg-12">
    <div class="card">
        <div class="card-header bg-primary">
            <strong class="card-title text-light"><i class="fa fa-edit"></i> Edit Profil</strong>
        </div>
        <?php 
        if(isset($data)){
            $d['d'] = $data->row();
            $d['form']= 'edit';
        ?>
        <form action="<?=base_url()?>profil/editSave/<?=$d['d']->id?>" method="post" enctype="multipart/form-data" class="form-horizontal">
            <div class="card-body card-block">
                    <?=$this->load->view('contents/profil/form', $d,TRUE)?>
            </div>
            <div class="card-footer">
                <button type="submit" class="btn btn-success btn-sm"><i class="fa fa-upload"> </i>PERBARUI</button>
                <a href="<?=base_url()?>profil" class="btn btn-danger btn-sm"><i class="fa fa-backward"> </i> KEMBALI</a>
            </div>
        </form>
        <?php
        }
        ?>
    </div>
</div>
</div>
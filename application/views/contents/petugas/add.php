
<div class="row">

<div class="col-lg-12">
    <div class="card">
        <div class="card-header bg-primary">
                <strong class="card-title text-light"><i class="fa fa-plus"></i> Tambah Data Petugas PBB</strong>
            </div>

        <form action="<?=base_url()?>petugas/addSave" method="post" class="form-horizontal">
            <div class="card-body card-block">
                    <?=$this->load->view('contents/petugas/form', '',TRUE)?>
            </div>
            <div class="card-footer">
                <button type="submit" class="btn btn-success btn-sm"><i class="fa fa-plus"></i>TAMBAHKAN
                </button>
                <a href="<?=base_url()?>petugas" class="btn btn-danger btn-sm"><i class="fa fa-backward"> </i> BATAL</a>
            </div>
        </form>
    </div>
</div>
</div>
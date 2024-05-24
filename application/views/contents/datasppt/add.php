<!-- halaman Tambah data sppt -->
<div class="row">
<div class="col-lg-12">
    <div class="card">
        <div class="card-header">
            <strong><i class="fa fa-plus"></i> Tambah Data SPPT</strong>
        </div>
        <form action="<?=base_url()?>datasppt/addSave" method="post" class="form-horizontal">
            <!-- form tambah data sppt -->
            <div class="card-body card-block">
                <?=$this->load->view('contents/datasppt/form', '',TRUE)?>
            </div>
            <div class="card-footer">
                <!-- Tombol tambah data sppt -->
                <button type="submit" class="btn btn-success btn-sm"><i class="fa fa-plus"></i>TAMBAHKAN
                </button>
                <a href="<?=base_url()?>datasppt" class="btn btn-danger btn-sm"><i class="fa fa-backward"> </i> BATAL</a>
            </div>
        </form>
    </div>
</div>
</div>
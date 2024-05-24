<div class="row">

    <div class="col-lg-12">
        <div class="card">
            <div class="card-header bg-success">
                <strong class="card-title text-light">Tentang Desa</strong>
            </div>

        </div>
        <div class="card-body card-block">
            <form action="" method="post" enctype="multipart/form-data" class="form-horizontal">
                <?php
                if (isset($data)) {
                    $d['d'] = $data->row();
                    $d['form'] = 'index';
                ?>
                    <?= $this->load->view('contents/profil/form', $d, TRUE) ?>
                <?php
                }
                ?>
            </form>
        </div>
        <div class="card-footer">
            <a href="<?= base_url() ?>profil/editForm" class="btn btn-success btn-m"><i class="fa fa-edit"></i> EDIT</a>
        </div>
    </div>
</div>
</div>
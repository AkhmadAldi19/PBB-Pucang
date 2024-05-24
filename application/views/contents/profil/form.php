<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Nama Desa/Kelurahan</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="nama_desa" class="form-control" value="<?= $d->nama_desa ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("nama_desa") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kecamatan</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kec" class="form-control" value="<?= $d->kec ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kec") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kabupaten</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kab" class="form-control" value="<?= $d->kab ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kab") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Nama Kepala Desa</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kades" class="form-control" value="<?= $d->kades ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kades") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Alamat Kantor</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kantor" class="form-control" value="<?= $d->kantor ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kantor") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Email</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="email" class="form-control" value="<?= $d->email ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("email") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Website Desa</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="web" class="form-control" value="<?= $d->web ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("web") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kode Wilayah</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kode" class="form-control" value="<?= $d->kode ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kode") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">No. Rekening</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="rek" class="form-control" value="<?= $d->rek ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("rek") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kop Surat Baris 1</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kop1" class="form-control" value="<?= $d->kop1 ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kop1") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kop Surat Baris 2</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kop2" class="form-control" value="<?= $d->kop2 ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kop2") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kop Surat Baris 3</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text" name="kop3" class="form-control" value="<?= $d->kop3 ?>" <?= $form == 'index' ? 'readonly' : '' ?>>
        <small><?= '<span class="text-danger">' . form_error("kop3") . '</span>' ?></small>
    </div>
</div>
<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Logo</label>
    </div>
    <div class="col-12 col-md-9">
        <?php
        if ($form == 'index') {
        ?>
            <img width="150px" src="<?= site_url() ?>assets/images/logo pucangsari.png" alt="">

        <?php
        } else {
        ?>
            <input type="file" id="file-input" name="logo" class="form-control-file">
            <small><?= '<span class="text-danger">' . form_error("logo") . '</span>' ?></small>
        <?php
        }
        ?>
    </div>
</div>
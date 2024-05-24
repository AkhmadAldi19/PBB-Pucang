<!-- Halaman Form edit dan tambah data sppt -->

    <!-- Form Untuk Tabel NOP -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">NOP</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="number"  name="nop" class="form-control" placeholder="Masukkan Nomor Objek Pajak" value="<?=!empty($d->nop)?$d->nop:''?>">
        </div>
    </div>

    <!-- Form Untuk Tabel Nama -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Nama</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="text"  name="nama" class="form-control" placeholder="Masukkan Nama" value="<?=!empty($d->nama)?$d->nama:''?>">
        </div>
    </div>

    <!-- Form Untuk Tabel Luas Bumi/Tanah -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Luas Bumi/Tanah</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="number"  name="bumi" class="form-control" placeholder="Masukkan Luas Bumi/Tanah" value="<?=!empty($d->bumi)?$d->bumi:''?>">
        </div>
    </div>

    <!-- Form Untuk Tabel Luas Bangunan -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Luas Bangunan</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="number"  name="bangunan" class="form-control" placeholder="Masukkan Luas Bangunan" value="<?=!empty($d->bangunan)?$d->bangunan:''?>">
        </div>
    </div>

    <!-- Form Untuk Tabel Jumlah Pajak -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Jumlah Pajak</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="number"  name="pajak" class="form-control" placeholder="Masukkan Jumlah Pajak" value="<?=!empty($d->pajak)?$d->pajak:''?>">
        </div>
    </div>

    <!-- Form Untuk Alamat Objek Pajak -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Alamat Objek Pajak</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="text"  name="alamat_op" class="form-control" placeholder="Masukkan Alamat Objek Pajak" value="<?=!empty($d->alamat_op)?$d->alamat_op:''?>">
        </div>
    </div>

    <!-- Form Untuk Alamat Wajib Pajak -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Alamat Wajib Pajak</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="text"  name="alamat_wp" class="form-control" placeholder="Masukkan Alamat Wajib Pajak" value="<?=!empty($d->alamat_wp)?$d->alamat_wp:''?>">
        </div>
    </div>

    <!-- Form Untuk Nomor HP -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="nomor_hp" class="form-control-label">Nomor HP</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="tel" id="nomor_hp" name="nomor_hp" class="form-control" value="<?=!empty($d->nomor_hp)?$d->nomor_hp:''?>" placeholder="Masukkan nomor HP" required>
        </div>
    </div>

    <!-- Form Untuk Nama Petugas -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="nomor_hp" class="form-control-label">Nama Petugas</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="text" id="nama_petugas" name="nama_petugas" class="form-control" value="<?=!empty($d->nama_petugas)?$d->nama_petugas:''?>" placeholder="Masukkan Nama Petugas" required>
        </div>
    </div>


    <!-- Form Keterangan -->
    <div class="row form-group">
        <div class="col col-md-3">
            <label for="text-input" class=" form-control-label">Keterangan</label>
        </div>
        <div class="col-12 col-md-9">
            <input type="text"  name="ket" class="form-control" value="Belum Bayar" readonly>
        </div>
    </div>




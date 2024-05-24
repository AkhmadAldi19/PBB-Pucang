<!-- Halaman Bayar Pajak -->
<div class="row">
    <div class="col-lg-12">
        <div class="card">
            <div class="card-header">
                <strong><i class="fa fa-money"></i> Bayar Pajak</strong>
            </div>

            <?php
            if(isset($data)){
                $d = $data->row();
            ?>

            <!-- Form Bayar Pajak -->
            <form action="<?=base_url()?>datasppt/bayarPajakSave/<?=$d->id?>" method="post" class="form-horizontal">
                <div class="card-body card-block">
                    <!-- NOP -->
                    <div class="row form-group">
                        <div class="col col-md-3">
                            <label for="text-input" class=" form-control-label">NOP</label>
                        </div>
                        <div class="col-12 col-md-9">
                            <input type="number"  name="nop" class="form-control" value="<?=$d->nop?>" readonly>
                        </div>
                    </div>

                    <!-- Nama -->
                    <div class="row form-group">
                        <div class="col col-md-3">
                            <label for="text-input" class=" form-control-label">Nama</label>
                        </div>
                        <div class="col-12 col-md-9">
                            <input type="text"  name="nama" class="form-control" value="<?=$d->nama?>" readonly>
                        </div>
                    </div>

                    <!-- Nilai Pajak PBB -->
                    <div class="row form-group">
                        <div class="col col-md-3">
                            <label for="text-input" class=" form-control-label">Nilai Pajak PBB</label>
                        </div>
                        <div class="col-12 col-md-9">
                            <input type="number"  name="pajak" class="form-control" value="<?=$d->pajak?>">
                        </div>
                    </div>
                    
                    <!-- Denda Pajak -->
                    <div class="row form-group">
                        <div class="col col-md-3">
                            <label for="text-input" class=" form-control-label">Denda Pajak</label>
                        </div>
                        <div class="col-12 col-md-9">
                            <input type="number"  name="denda" class="form-control" value="0">
                        </div>
                    </div>
                </div>

                <div class="card-footer">
                    <!-- Tombol Submit -->
                    <button type="submit" class="btn btn-success btn-sm">
                        SUBMIT
                    </button>
                </div>
            
            </form>
            <?php } ?>
        </div>
    </div>
</div>
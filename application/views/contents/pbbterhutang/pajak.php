
<div class="row">

<div class="col-lg-12">
    <div class="card">
        <div class="card-header">
            <strong><i class="fa fa-money"></i> Bayar PBB</strong>
        </div>
        <?php 
        
        if(isset($data)){
            $d = $data->row();
        ?>
        <form action="<?=base_url()?>pbbterhutang/bayarPajakSave/<?=$d->id?>" method="post" class="form-horizontal">
            <div class="card-body card-block">
                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">NOP</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="number"  name="nop" class="form-control" value="<?=$d->nop?>" readonly>
                    </div>
                </div>

                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">Nama</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="text"  name="nama" class="form-control" value="<?=$d->nama?>" readonly>
                    </div>
                </div>
                
                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">Nilai Pajak PBB</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="number"  name="pajak" class="form-control" value="<?=$d->pajak?>">
                    </div>
                </div>
<!-- 
                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">Nilai Pajak pada SKP</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="number"  name="skp" class="form-control" >
                    </div>
                </div>

                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">Nilai Pajak pada STP</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="number"  name="stp" class="form-control" >
                    </div>
                </div> -->
                <div class="row form-group">
                    <div class="col col-md-3">
                        <label for="text-input" class=" form-control-label">Periode Pajak</label>
                    </div>
                    <div class="col-12 col-md-9">
                        <input type="number"  name="tahun" class="form-control" placeholder="Isi Periode Tahun Pajak ">
                    </div>
                </div>
                
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
                <button type="submit" class="btn btn-success btn-sm"><i class="fa fa-money"></i> BAYAR
                </button>
                <a href="<?=base_url()?>pbbterhutang" class="btn btn-danger btn-sm"><i class="fa fa-backward"> </i> BATAL BAYAR</a>
            </div>
            </div>
        </form>
        <?php } ?>
    </div>
</div>
</div>
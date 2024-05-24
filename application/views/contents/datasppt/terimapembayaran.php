<!-- Halaman Terima Pembayaran -->
<div class="row">
    <div class="col-lg-12">
        <div class="card">
            <div class="card-header ">
                <strong>Terima Pembayaran</strong>
            </div>
            <div class="card-body card-block">

                <!-- Memilih Pembayar -->
                <div class="row form-group">
                    <div class="col-12 col-md-9">

                        <select name="select" id="select" class="selectpicker form-control" data-container="body" data-live-search="true" required style="background-color: green;">
                            <option value="0">Please select</option>
                            <?php
                            if ($datas->num_rows() > 0) {
                                foreach ($datas->result() as $d) {
                            ?>
                                    <option value="<?= $d->id ?>"><?= $d->nop . ' - ' . $d->nama ?></option>
                            <?php
                                }
                            }
                            ?>
                        </select>
                        
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
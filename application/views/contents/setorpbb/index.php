
<div class="row">
    <div class="col-lg-12">
    
        <form action="<?=base_url()?>pbblunas/setor" method="post">
            <!-- USER DATA-->
            <div class="user-data m-b-30">
                <h3 class="title-3 m-b-30 ml-2">
                    Data PBB Setor</h3>
                    
                    <div class="table-data__tool">



                        <div class="table-data__tool-right ml-5">
                            <div class="row">
                            
                                <div class="col-6">
                                    <div class="form-group">
                                        <label for="cc-exp" class="control-label mb-1">Tanggal Awal</label>
                                        <input id="tgl_awal" name="tgl_awal" type="date" class="form-control cc-exp">
                                    </div>
                                </div>
                                <div class="col-6">
                                    <div class="form-group">
                                        <label for="cc-exp" class="control-label mb-1">Tanggal Akhir</label>
                                        <input id="tgl_akhir" name="tgl_akhir" type="date" class="form-control cc-exp">
                                    </div>
                                </div>
                            </div>
                            <a href="#" id="cetaksetor" class="au-btn au-btn-icon au-btn--blue au-btn--small text-white">
                            <i class="fa fa-print"></i>Cetak Laporan</a>
                        </div>
                    </div>

                <?php 
                if($this->uri->segment(2) == 'setored'){
                ?>

                <div class="alert alert-primary" role="alert">
                    Data berhasil di setorkan
                </div>
                <?php

                }
                ?>
                <?php 
                if($this->uri->segment(2) == 'bataled'){
                ?>

                <div class="alert alert-primary" role="alert">
                    Data telah dibatalkan
                </div>
                <?php

                }
                ?>
                <div class="table-responsive ">
                <table class="table table-borderless table-data3" id="myTable">
                        <thead>
                            <tr>
                                <th>No</th>
                                <th>NOP</th>
                                <th>Nama WP</th>
                                <th>Tahun</th>
                                <th>Pajak (Rp.)</th>
                            </tr>
                        </thead>
                        <tbody>
                        <?php 
                        if($datas->num_rows() >0){
                            $no = 0;
                            foreach($datas->result() as $d){
                                $no++;
                        ?>
                            <tr>
                                <td><?=$no?></td>
                                <td><?=$d->nop?></td>
                                <td><?=$d->nama?></td>
                                <td><?=$d->tahun?></td>
                                <td><?=$d->total_pajak?></td>
                            </tr>
                        <?php
                            }
                        }
                        ?>
                        </tbody>
                    </table>
                </div>
            </div>
            <!-- END USER DATA-->
        </form>
    </div>
</div>

<div class="row mb-4">
    <div class="col">
    <?=$pagination?>
    </div>
</div>
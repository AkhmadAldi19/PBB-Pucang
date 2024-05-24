
<div class="row">
    <div class="col-lg-12">
        <!-- USER DATA-->
        <div class="user-data m-b-30">
            <h3 class="title-3 m-b-30">
                PBB Terhutang
                <a href="<?=base_url()?>cetakdokumen/index" target="_blank" data-toggle="tooltip" data-placement="top" title="Cetak Laporan"><i class="fa fa-print"></i> </a></h3>
                

            <?php 
            if($this->uri->segment(2) == 'added'){
            ?>

            <div class="alert alert-primary" role="alert">
                Data berhasil di tambah
            </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'edited'){
            ?>

            <div class="alert alert-primary" role="alert">
                Data berhasil di edit
            </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'deleted'){
            ?>

            <div class="alert alert-primary" role="alert">
                Data berhasil di hapus
            </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'bayar'){
            ?>

            <div class="alert alert-primary" role="alert">
                Sudah bayar cek pada laporan PBB lunas untuk setor
            </div>
            <?php

            }
            ?>
            <div class="table-responsive ">
                <table class="table table-borderless table-data3" id="myTable">
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Aksi</th>
                            <th>Nama/NOP/Alamat WP/Objek/Luas</th>
                            <th>Pajak (Rp.)</th>
                            <th>Keterangan</th>
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
                            <td>
                            <a href="<?=base_url()?>pbbterhutang/bayarPajak/<?=$d->id?>" class="text-primary lead" data-toggle="tooltip" data-placement="top" title="Bayar Pajak"><i class="fa fa-money"></i></a>
                            <!--<a href="#" class="text-primary lead" data-toggle="tooltip" data-placement="top" title="Edit Data"><i class="fa fa-edit"></i></a>
                            <a href="#" class="text-primary lead" data-toggle="tooltip" data-placement="top" title="Edit Data"><i class="fa fa-edit"></i></a> -->
                            </td>
                            <td>Nama : <?=$d->nama?> 
                            <br />
                            NOP : <?=$d->nop?>
                            <br />
                            Alamat WP : <?=$d->alamat_wp?>
                            <br />
                            Alamat Objek : <?=$d->alamat_op?>
                            <br />
                            Luas Bumi: <?=$d->bumi?>
                            Bangunan: <?=$d->bangunan?></td>
                            <td><?=$d->pajak?></td>
                            <td><?=$d->ket?></td>
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
    </div>
</div>

<div class="row mb-4">
    <div class="col">
    <?=$pagination?>
    </div>
</div>
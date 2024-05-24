<!-- Halaman Index data Sppt -->
<div class="row">
    <div class="col-lg-12">
        <!-- USER DATA-->
        <div class="user-data m-b-30">
            <h3 class="title-3 m-b-30">
                Data SPPT
            </h3>
                <!-- Tombol Tambah data dan import DHKP -->
                <div class="table-data__tool ml-5">
                    <div class="table-data__tool-left">
                        <a href="<?=base_url()?>datasppt/addForm" class="au-btn au-btn-icon au-btn--green au-btn--small text-white">
                            <i class="zmdi zmdi-plus"></i>Tambah Data</a>
                        <a href="<?=base_url()?>Welcome" class="au-btn au-btn-icon au-btn--blue au-btn--small text-white">
                            <i class="zmdi zmdi-database"></i>Import DHKP</a>
                    </div>
                </div>
            <!-- beberapa peringatan -->
            <?php
            if($this->uri->segment(2) == 'added'){
            ?>
                <div class="alert alert-primary" role="alert">
                    Yey Data berhasil di tambah 👌
                </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'edited'){
            ?>
                <div class="alert alert-primary" role="alert">
                    Yey Data berhasil di Edit 👌
                </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'deleted'){
            ?>
                <div class="alert alert-primary" role="alert">
                    Data Sukses di hapus ni 😊
                </div>
            <?php

            }
            ?>
            <?php 
            if($this->uri->segment(2) == 'bayar'){
            ?>
                <div class="alert alert-primary" role="alert">
                    Sudah di bayar nih yuk cek pada laporan PBB lunas untuk setor 👌
                </div>
            <?php

            }
            ?>

            <!-- Tabel -->
            <div class="table-responsive ">
                <table class="table table-borderless table-data3" id="myTable">
                    <!-- Kepala Tabel -->
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Aksi</th>
                            <th>Nama/NOP/Alamat WP/Objek/Luas</th>
                            <th>Pajak</th>
                            <th>Keterangan</th>
                        </tr>
                    </thead>
                    <!-- Body tabel -->
                    <tbody>
                        <?php 
                            if($datas->num_rows() >0){
                                $no = 0;
                                foreach($datas->result() as $d){
                                    $no++;
                            ?>
                            <tr>
                                <!-- isi tabel Nomor -->
                                <td>
                                    <?=$no?>
                                </td>
                                <!-- isi tabel Aksi -->
                                <td>
                                    <!-- edit -->
                                    <a href="<?=base_url()?>datasppt/editForm/<?=$d->id?>" class="text-primary lead" data-toggle="tooltip" data-placement="top" title="Edit Data"><i class="fa fa-edit">
                                        </i>
                                    </a>
                                    
                                    <!-- Delete -->
                                    <a href="<?=base_url()?>datasppt/deleteData/<?=$d->id?>" class="text-danger lead" data-toggle="tooltip" data-placement="top" title="Hapus Data" OnClick="return confirm('Apa anda yakin ingin mendelete ini?');"><i class="fa fa-trash">
                                        </i>
                                    </a>

                                    <!-- Jika Belum Membayar muncul tombol bayar pajak  -->
                                    <?php if($d->ket == 'Belum Bayar') { ?>
                                        <!-- Tambahkan tautan WhatsApp di sini -->
                                        <a href="https://wa.me/<?=$d->nomor_hp?>?text=Halo%20<?=$d->nama?>,%20ini%20pesan%20peringatan%20otomatis%20dari%20sistem%20kami.%20Terima%20kasih%20%20 <?=base_url()?>cetakdokumen/tandaTerima/<?=$d->id?>" class="text-info lead ml-2" data-toggle="tooltip" data-placement="top" title="Kirim Pesan WhatsApp">
                                            <i class="fab fa-whatsapp"></i>
                                        </a>
                                        
                                        <a href="<?=base_url()?>datasppt/bayarPajak/<?=$d->id?>" class="text-info lead" data-toggle="tooltip" data-placement="top" title="Bayar Pajak"><i class="fa fa-money">
                                            </i>
                                        </a>
                                        
                                    <!-- Jika Sudah lunas muncul tombol whatsapp  -->
                                    <?php } elseif($d->ket == 'Lunas') { ?>
                                        <a href="https://wa.me/<?=$d->nomor_hp?>?text=Halo%20<?=$d->nama?>,%20ini%20pesan%20otomatis%20dari%20sistem%20kami.%20Terima%20kasih%20%20 <?=base_url()?>cetakdokumen/tandaTerima/<?=$d->id?>" class="text-info lead" data-toggle="tooltip" data-placement="top" title="Kirim Pesan WhatsApp"><i class="fab fa-whatsapp">
                                            </i>
                                        </a>
                                    <?php } ?>

                                    <!-- Tanda Terima -->
                                    <a href="<?=base_url()?>cetakdokumen/tandaTerima/<?=$d->id?>" target="_blank"class="text-secondary lead" data-toggle="tooltip" data-placement="top" title="Tanda Terima"><i class="fa fa-print">
                                        </i>
                                    </a>
                                </td>

                                <!-- Nama/NOP/Alamat WP/Objek/Luas -->
                                <td>
                                    Nama : <?=$d->nama?> 
                                    <br />
                                    NOP : <?=$d->nop?>
                                    <br />
                                    Alamat WP : <?=$d->alamat_wp?>
                                    <br />
                                    Alamat Objek : <?=$d->alamat_op?>
                                    <br />
                                    Luas Bumi: <?=$d->bumi?>
                                    <br />
                                    Bangunan: <?=$d->bangunan?>
                                    <br/>
                                    nomor hp: <?=$d->nomor_hp?>
                                    <br/>
                                    Nama Petugas: <?=$d->nama_petugas?>
                                </td>

                                <!-- Pajak -->
                                <td><?=$d->pajak?></td>
                                <!-- Keterangan -->
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


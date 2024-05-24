
<div class="row">
    <div class="col-lg-12">
    
        <form action="<?=base_url()?>pbblunas/setor" method="post">
            <!-- USER DATA-->
            <div class="user-data m-b-30">
                <h3 class="title-3 m-b-30">
                    Data PBB Lunas 
                    <a href="<?=base_url()?>cetakdokumen/pbblunas" target="_blank" data-toggle="tooltip" data-placement="top" title="Cetak Laporan"><i class="fa fa-print"></i> </a></h3>
                    
                    <div class="table-data__tool">
                        <div class="table-data__tool-left">
                            <div class="rs-select2--light rs-select2--sm">
                                <select class="js-select2" name="type">
                                    <option selected="selected" value="setor">Setor</option>
                                    <option value="batal">Batal Setor</option>
                                </select>
                                <div class="dropDownSelect2"></div>
                            </div>
                            
                            <button id="sendData" type="submit" class="au-btn au-btn-icon au-btn--green au-btn--small text-white">Submit</button>
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
                    <table class="table table-borderless table-data2" id="myTable">
                        <thead>
                            <tr>
                                <th>
                                    <label class="au-checkbox">
                                        <input type="checkbox" id="allcheck">
                                        <span class="au-checkmark"></span>
                                    </label>
                                </th>
                                <th>No</th>
                                <th>NOP</th>
                                <th>Tahun</th>
                                <th>Nama WP</th>
                                <th>Pajak (Rp.)</th>
                                <th>Keterangan</th>
                                <th>Petugas</th>

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
                                <td>
                                    <label class="au-checkbox">
                                        <input type="checkbox" class="check" name="collection[id][]" value="<?=$d->id . ';' . $d->nop . ';' . $d->pajak?>">
                                        <span class="au-checkmark"></span>
                                    </label>
                                </td>
                                <td><?=$no?></td>
                                <td><?=$d->nop?></td>
                                <td><?=$d->tahun?></td>
                                <td><?=$d->nama?></td>
                                <td><?=$d->pajak?></td>
                                <td><?=$d->ket?></td>
                                <td><?=$d->nm_petugas?></td>
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
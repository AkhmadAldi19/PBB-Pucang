
<div class="row">
    <div class="col-lg-12">
        <!-- USER DATA-->
        <div class="user-data m-b-30">
            <h3 class="title-3 m-b-20 m-l-5">
                PETUGAS</h3>
                <div class="table-data__tool ml-5">
                    <div class="table-data__tool-left">
                            <a href="<?=base_url()?>petugas/addForm" class="au-btn au-btn-icon au-btn--green au-btn--small text-white">
                            <i class="zmdi zmdi-plus"></i>Tambah Data</a>
                    </div>
                </div>



                 

            <div class="table-responsive ">
                <table class="table table-borderless table-data3" id="myTable">
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Username</th>
                            <th>No.Telp</th>
                            <th>Level</th>
                            <th>Aksi</th>
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
                            <td><?=$d->uname?></td>
                            <td><?=$d->hp?></td>
                            <td><?=$d->level?></td>
                            <td>
                                <a href="<?=base_url()?>petugas/editForm/<?=$d->id?>" class="text-primary lead" data-toggle="tooltip" data-placement="top" title="Edit Data"><i class="fa fa-edit"></i></a>

                            <a href="<?=base_url()?>petugas/deleteData/<?=$d->id?>" class="text-danger lead" data-toggle="tooltip" data-placement="top" title="Hapus Data" OnClick="return confirm('Apa anda yakin ingin mendelete ini?');"><i class="fa fa-trash"></i></a></td>
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
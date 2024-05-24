                    <?php

                    if (isset($data)) {
                        $d = $data->row();
                    ?>


                        <div class="container-fluid">
                            <div class="row">
                                <div class="col-md-12">
                                    <div class="overview-wrap">
                                        <h2 class="title-1">Dashboard</h2>

                                    </div>
                                </div>
                            </div>
                            <div class="row m-t-25">

                                <!-- Jumlah nop -->
                                <div class="col-sm-6 col-lg-12">
                                    <div class="overview-item overview-item--c4">
                                        <div class="overview__inner">
                                            <div class="overview-box clearfix p-3">
                                                <div class="icon">
                                                    <i class="zmdi zmdi-account-o"></i>
                                                </div>
                                                <div class="text">
                                                    <span class="desc" style="font-weight: 600; color: #fff; font-size: 22px ;"> Jumlah NOP <?= isset($d->jumlah_nop) ? $d->jumlah_nop : 'N/A' ?></span>
                                                    <h3 class="number" style="font-size: 16px; font-weight: 400; color: #fff;"> PBB =
                                                        <?= isset($d->total_pajak) ? $rupiah($d->total_pajak) : 'N/A' ?>
                                                    </h3>
                                                </div>

                                                <h3 class="desc m-b-10 m-t-5" style="font-size: 16px; font-weight: 300; color: #fff ; ">
                                                NOP adalah singkatan dari Nomor Objek Pajak yang memiliki arti susunan nomor identitas objek pajak.</h3>

                                                <button style="font-weight: 600;" class=" custom-button m-t-5 " onclick="window.location='<?= base_url() ?>datasppt'">
                                                    Lihat Data
                                                </button>
                                            </div>
                                            <hr>
                                            <div class="overview-chart">
                                                <canvas id="widgetChart1"></canvas>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- end -->

                                <!-- PBB Lunas -->
                                <div class="col-sm-6 col-lg-12">
                                    <div class="overview-item overview-item--c3">
                                        <div class="overview__inner">
                                            <div class="overview-box clearfix p-3">
                                                <div class="icon">
                                                    <i class="zmdi zmdi-card"></i>
                                                </div>
                                                <div class="text">
                                                    <span class="desc" style="font-weight: 600; color: #fff; font-size: 22px ;">Lunas <?= isset($d->jumlah_nop) ? $d->lunas : 'N/A' ?></span>
                                                    <h3 class="number" style="font-size: 16px; font-weight: 400; color: #fff;"> PBB =
                                                    <?php if (isset($d->pajak_lunas)) {
                                                        echo $rupiah($d->pajak_lunas);
                                                    }
                                                    ?>
                                                    </h3>
                                                </div>

                                                <h3 class="desc m-b-10 m-t-5" style="font-size: 16px; font-weight: 300; color: #fff ; ">
                                                    Jumlah NOP Yang Terhitung Lunas diambil dari jumlah NOP</h3>

                                                <button style="font-weight: 600;" class=" custom-button-red m-t-5 " onclick="window.location='<?= base_url() ?>pbblunas'">
                                                    Lihat Data
                                                </button>
                                            </div>
                                            <hr>
                                            <div class="overview-chart">
                                                <canvas id="widgetChart2"></canvas>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- end -->

                                <!-- PBB Terhutang -->
                                <div class="col-sm-6 col-lg-12">
                                    <div class="overview-item overview-item--c4">
                                        <div class="overview__inner">
                                            <div class="overview-box clearfix p-3">
                                                <div class="icon">
                                                    <i class="zmdi zmdi-balance-wallet"></i>
                                                </div>
                                                <div class="text">
                                                    <span class="desc" style="font-size: 22px; font-weight: 600; color: #fff;">PBB Terhutang <?= isset($d->jumlah_nop) ? $d->terhutang : 'N/A' ?></span>
                                                    <h3 class="number" style="font-size: 16px; font-weight: 400; color: #fff;">
                                                        <?php
                                                        if ($d->pajak_terhutang !== null) {
                                                            echo $rupiah($d->pajak_terhutang);
                                                        } else {
                                                            echo 'N/A';
                                                        }
                                                        ?>
                                                    </h3>
                                                </div>

                                                <h3 class="desc m-b-10 m-t-5" style="font-size: 16px; font-weight: 300; color: #fff ; ">
                                                Jumlah NOP Yang Terhitung Terhutang diambil dari jumlah NOP</h3>

                                                <button style="font-weight: 600;" class=" custom-button m-t-5 " onclick="window.location='<?= base_url() ?>pbbterhutang'">
                                                    Lihat Data
                                                </button>
                                            </div>
                                            <hr>
                                            <div class="overview-chart">
                                                <canvas id="widgetChart3"></canvas>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- end -->

                                <!-- Presentase Lunas -->
                                <div class="col-sm-6 col-lg-12">
                                    <div class="overview-item overview-item--c3">
                                        <div class="overview__inner">
                                            <div class="overview-box clearfix p-3">
                                                <div class="icon">
                                                    <i class="zmdi zmdi-chart-donut"></i>
                                                </div>
                                                <div class="text">
                                                    <span class="desc" style="font-size: 22px; font-weight: 600; color: #fff;">Presentase Lunas</span>
                                                    <h3 class="number" style="font-size: 16px; font-weight: 400; color: #fff;">
                                                    <?= number_format($d->presentase, 0) ?> %
                                                    </h3>
                                                </div>

                                                <h3 class="desc m-b-10 m-t-5" style="font-size: 16px; font-weight: 300; color: #fff ; ">
                                                    Jumlah Lunas NOP Yang Terhitung Lunas dituli dalam bentuk persen</h3>

                                                <button style="font-weight: 600;" class=" custom-button-red m-t-5 " onclick="window.location='<?= base_url() ?>datasppt'">
                                                    Lihat Data
                                                </button>
                                            </div>
                                            <hr>
                                            <div class="overview-chart">
                                                <canvas id="widgetChart4"></canvas>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- end -->

                            </div>
                            <!-- STATISTIC-->
                            <section class="statistic statistic2">
                                <div class="container">
                                    <div class="row">
                                        <div class="col-md-6 col-lg-3">
                                            <div class="statistic__item statistic__item--green">
                                                <?php if ($d->jumlah_nop !== null) : ?>
                                                    <h2 class="number" style="font-weight: 500; color: #fff; font-size: 28px ;"><?= $d->jumlah_nop ?></h2>
                                                <?php endif; ?>
                                                <span class="desc" style="font-weight: 600; color: #fff; font-size: 18px ;">Wajib Pajak</span>
                                                <div class="icon">
                                                    <i class="zmdi zmdi-account-o"></i>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="col-md-6 col-lg-3">
                                            <div class="statistic__item statistic__item--orange">
                                                <?php if ($d->jumlah_nop !== null) : ?>
                                                    <h2 class="number" style="font-weight: 500; color: #fff; font-size: 28px ;"><?= $d->terhutang ?></h2>
                                                <?php endif; ?>
                                                <span class="desc" style="font-weight: 600; color: #fff; font-size: 18px ;">WP Belum Bayar</span>
                                                <div class="icon">
                                                    <i class="zmdi zmdi-account-o"></i>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="col-md-6 col-lg-3">
                                            <div class="statistic__item statistic__item--blue">
                                                <?php if ($d->jumlah_nop !== null) : ?>
                                                    <h2 class="number" style="font-weight: 500; color: #fff; font-size: 28px ;"><?= $d->lunas ?></h2>
                                                <?php endif; ?>
                                                <span class="desc" style="font-weight: 600; color: #fff; font-size: 18px ;">WP Lunas</span>
                                                <div class="icon">
                                                    <i class="zmdi zmdi-account-o"></i>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="col-md-6 col-lg-3">
                                            <div class="statistic__item statistic__item--red">
                                                <?php if ($d->jumlah_nop !== null) : ?>
                                                    <h2 class="number" style="font-weight: 500; color: #fff; font-size: 28px ;"><?= number_format($d->presentase, 0) ?> %
                                                <?php endif; ?>
                                                <br>
                                                <span class="desc" style="font-weight: 600; color: #fff; font-size: 18px ;">PBB Masuk</span>
                                                <div class="icon">
                                                    <i class="zmdi zmdi-account-o"></i>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </section>
                            <!-- END STATISTIC-->

                        <?php

                    }
                        ?>
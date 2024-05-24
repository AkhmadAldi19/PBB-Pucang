<!-- HEADER Mobile-->
<header class="header-desktop2">
    <div class="section__content section__content--p30">
        <div class="container-fluid">
            <div class="header-wrap2">
                <div class="logo d-block d-lg-none">
                    <a href="<?= base_url() ?>dashboard">
                        <img src="<?= site_url() ?>assets/images/admin.png" alt="Admin PBB Desa Pucangsari" width="140px" height="auto" style="margin-right: 15px; margin-top: 10px; margin-bottom: 10px;" />
                    </a>
                </div>
                <!-- <marquee style="font-family: Arial; font-size:12px; color:white; ">Selamat Datang di Sistem Informasi dan Aplikasi Pencatatan dan Pelunasan PBB Desa Pucangsari Kecamatan Purwosari </marquee> -->
                <div class="header-button2">
                    <div class="header-button-item js-item-menu">
                        <i class="zmdi zmdi-info"><a href=""></i>
                        </a>
                    </div>
                    <div class="header-button-item mr-0 js-sidebar-btn">
                        <i class="zmdi zmdi-menu"></i>
                    </div>
                    <div class="setting-menu js-right-sidebar d-none d-lg-block">
                        <div class="account-dropdown__body">
                            <div class="account-dropdown__item">
                                <a href="#">
                                    <i class="zmdi zmdi-account"></i>Profil</a>
                            </div>
                            <div class="account-dropdown__item">
                                <a href="#">
                                    <i class="zmdi zmdi-settings"></i>Pengaturan</a>
                            </div>
                            <div class="account-dropdown__item">
                                <a href="<?= base_url() ?>login/logout">
                                    <i class="zmdi zmdi-power"></i>Keluar Sistem</a>
                            </div>
                        </div>

                    </div>
                </div>
            </div>
        </div>
    </div>
</header>
<aside class="menu-sidebar2 js-right-sidebar d-block d-lg-none">
    <div class="logo">
        <!-- <a href="#">
            <img src="<?= site_url() ?>assets/images/logo pasuruan.png" alt="Logo Siap Ayem 2" />
        </a> -->
    </div>
    <div class="menu-sidebar2__content js-scrollbar2">
        <div class="account2">
            <div class="image img-cir img-120">
                <img src="<?= site_url() ?>assets/images/profile.jpg" alt="User Foto" />
            </div>
            <h4 class="name"><a style="color:#009A4C; text-align: center; font-size: 20px; font-weight: 500; text-transform: uppercase; " class=" js-acc-btn" href="#"><?= $this->session->userdata('username') ?></a>
            </h4> <a href="#" style="text-align: center; font-size:12px; font-weight: 500; color: #303030;"><?= $this->session->userdata('hp') ?></a></h4>
            <!-- <a href="<?= base_url() ?>login/logout">Keluar </a> -->
        </div>
        <nav class="navbar-sidebar2">
            <ul class="list-unstyled navbar__list">
                <li class="active has-sub">
                    <a class="js-arrow" href="<?= base_url() ?>dashboard">
                        <i class="fas fa-tachometer-alt"></i>Dashboard
                    </a>
                </li>
                <!-- Terima Pembayaran -->
                <li>
                    <a class="js-arrow" href="<?= base_url() ?>datasppt/terimapembayaran">
                        <i class="fa fa-money"></i>Terima Pembayaran
                    </a>
                </li>
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-th"></i>Master Data <i class="fa fa-angle-left pull-right"></i></a>
                    <ul class="list-unstyled navbar__sub-list js-sub-list">
                        <li><a href="<?= base_url() ?>profil"><i class="fas fa-map-marker-alt"></i> Profil Desa</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>datasppt"><i class="fas fa-database"></i> Data DHKP</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>petugas"><i class="fas fa-user"></i> Data Petugas</a>
                        </li>
                    </ul>
                </li>
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-print"></i>Laporan <i class="fa fa-angle-left pull-right"></i></a>
                    <ul class="list-unstyled navbar__sub-list js-sub-list">
                        <li>
                            <a href="<?= base_url() ?>pbbterhutang"><i class="fas fa-circle"></i> PBB Terhutang</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>pbblunas"><i class="fas fa-circle"></i> PBB Lunas</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>setorpbb"><i class="fas fa-circle"></i> Setor PBB</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>rekapitulasipbb"><i class="fas fa-circle"></i> Rekapitulasi PBB</a>
                        </li>
                    </ul>
                </li>

                <li>
                    <a href="<?= base_url() ?>login/logout"><i class="zmdi zmdi-power"></i> Keluar Sistem</a>
                </li>
            </ul>
            </li>
            </ul>
            </li>
            </ul>
        </nav>
    </div>
</aside>
<!-- END HEADER DESKTOP-->
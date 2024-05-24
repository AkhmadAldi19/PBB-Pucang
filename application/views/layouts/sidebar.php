<!-- MENU SIDEBAR-->
<aside class="menu-sidebar2">
    <!-- Logo Sidebar Header Dekstop -->
    <div class="logo">
        <a href="<?= base_url() ?>dashboard">
            <div style="display: flex; align-items: center;">
                <img src="<?= site_url() ?>assets/images/admin.png" alt="admin pbb desa Pucangsari" width="180px" height="auto" style="margin-right: 15px; margin-top: 10px; margin-bottom: 10px;" />
                <!-- <span style="color: white; font-size: 23px;">ADMIN<br>PBB Desa</span> -->
            </div>
        </a>
    </div>

    <!-- Menu Sidebar Header Dekstop  -->
    <div class="menu-sidebar2__content js-scrollbar1">
        <!-- <h3 class="name" style="text-align: center; font-size:12px; margin-top: 60px;">
            <a style="text-align: center; font-size:16px " class="js-acc-btn" href="#"><?= $this->session->userdata('username') ?></a>
        </h3> -->

        <!-- Welcome Dashboard -->
        <div class="account2">
            <!-- Foto Profile -->
            <div class="image img-cir img-12">
                <img src="<?= site_url() ?>assets/images/profile.jpg" alt="User Foto" style="width: 100px; height: 100px;" />
            </div>
            <!-- Username -->
            <h4>
                <a class="name" style="color:#009A4C; text-align: center; font-size: 24px; font-weight: 500; text-transform: uppercase; " href="#">
                    <?= $this->session->userdata('username') ?></a>
            </h4>
            <!-- No Handphone -->
            <h4 class="hp" class="name" style="text-align: center; font-size:14px; font-weight: 500; color: #303030;"><?= $this->session->userdata('hp') ?></h4>
        </div>

        <!-- Main Menu -->
        <nav class="navbar-sidebar2">
            <ul class="list-unstyled navbar__list">
                <!-- Dashboard -->
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
                <!-- DATA SPPT -->
                <li>
                    <a href="<?= base_url() ?>datasppt"><i class="fas fa-database"></i> DATA SPPT </a>
                </li>
                <!-- Pengaturan -->
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-cogs"></i>Pengaturan <i class="fa fa-angle-left pull-right"></i>
                    </a>
                    <!-- Dropdown Pengaturan -->
                    <ul class="list-unstyled navbar__sub-list js-sub-list">
                        <li>
                            <a href="<?= base_url() ?>profil"><i class="fas fa-map-marker-alt"></i> Profil Wilayah</a>
                        </li>
                        <li>
                            <a href="<?= base_url() ?>petugas"><i class="fas fa-user"></i> Manajemen Pengguna</a>
                        </li>
                    </ul>
                </li>
                <!-- Laporan -->
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-print"></i>Laporan <i class="fa fa-angle-left pull-right"></i>
                    </a>
                    <!-- Dropdown Pengaturan -->
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
                <!-- Keluar Sistem -->
                <li>
                    <a href="<?= base_url() ?>login/logout"><i class="zmdi zmdi-power"></i> Keluar Sistem</a>
                </li>
            </ul>
            </li>
            </ul>
        </nav>
    </div>
</aside>
<!-- END MENU SIDEBAR-->
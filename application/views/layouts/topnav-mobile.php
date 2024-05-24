<!--HEADER MOBILE-->
<header class="header-mobile d-block d-lg-none">
    <div class="header-mobile__bar">
        <div class="container-fluid">
            <div class="header-mobile-inner">
                <a class="logo" href="index.html">
                    <img src="<?= site_url() ?>assets/images/logo pasuruan.png" alt="CoolAdmin" />
                </a>
                <button class="hamburger hamburger--slider" type="button">
                    <span class="hamburger-box">
                        <!-- <span class="hamburger-inner"></span> -->
                    </span>
                </button>
            </div>
        </div>
    </div>
    <nav class="navbar-mobile">
        <div class="container-fluid">
            <ul class="navbar-mobile__list list-unstyled">
                <li>
                    <a class="js-arrow" href="<?= base_url() ?>dashboard">
                        <i class="fas fa-tachometer-alt"></i>Dashboard</a>
                </li>
                <li>
                    <a class="js-arrow" href="<?= base_url() ?>datasppt/terimapembayaran">
                        <i class="fa fa-money"></i>Terima Pembayaran
                    </a>
                    <a class="js-arrow" href="#">
                        <i class="fas fa-tachometer-alt"></i>Terima Pembayaran</a>
                </li>
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-copy"></i>Master Data <i class="fa fa-angle-left pull-right"></i></a>
                    <ul class="list-unstyled navbar__sub-list js-sub-list">
                        <li>
                            <a href="<?= base_url() ?>profil">Profil</a>
                        </li>
                        <li>
                            <a href="register.html">Data SPPT</a>
                        </li>
                        <li>
                            <a href="forget-pass.html">Data Petugas</a>
                        </li>
                        <li>
                            <a href="forget-pass.html">Data Wajib Pajak</a>
                        </li>
                    </ul>
                </li>
                <li class="has-sub">
                    <a class="js-arrow" href="#">
                        <i class="fas fa-desktop"></i>Laporan <i class="fa fa-angle-left pull-right"></i></a>
                    <ul class="list-unstyled navbar__sub-list js-sub-list">
                        <li>
                            <a href="button.html">Pembarayan Online</a>
                        </li>
                        <li>
                            <a href="badge.html">PBB Lunas</a>
                        </li>
                        <li>
                            <a href="tab.html">PBB Terhutang</a>
                        </li>
                        <li>
                            <a href="card.html">Setor PBB</a>
                        </li>
                        <li>
                            <a href="alert.html">Rekapitulasi PBB</a>
                        </li>
                    </ul>
                </li>

                <li>
                    <a href="progress-bar.html">Logout</a>
                </li>
            </ul>
        </div>
    </nav>
</header>
<!-- END HEADER MOBILE
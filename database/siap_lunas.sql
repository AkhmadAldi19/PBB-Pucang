-- phpMyAdmin SQL Dump
-- version 5.1.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Jan 10, 2022 at 08:47 AM
-- Server version: 10.4.22-MariaDB
-- PHP Version: 8.0.13

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `siap_lunas`
--

-- --------------------------------------------------------

--
-- Table structure for table `admin`
--

CREATE TABLE `admin` (
  `id` int(11) NOT NULL,
  `uname` varchar(30) NOT NULL,
  `hp` varchar(100) NOT NULL,
  `pass` varchar(70) NOT NULL,
  `tgl` date NOT NULL,
  `foto` varchar(150) NOT NULL,
  `level` varchar(15) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

--
-- Dumping data for table `admin`
--

INSERT INTO `admin` (`id`, `uname`, `hp`, `pass`, `tgl`, `foto`, `level`) VALUES
(1, 'admindesa', '085647986900', 'b82948d4da194898c5672b49847115c6', '2022-01-09', '1.jpg', 'admin');

-- --------------------------------------------------------

--
-- Table structure for table `data_bayar`
--

CREATE TABLE `data_bayar` (
  `id` int(11) NOT NULL,
  `nop` varchar(20) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `pajak` int(12) NOT NULL,
  `tahun` int(12) NOT NULL,
  `stp` int(12) NOT NULL,
  `denda` int(12) NOT NULL,
  `total_pajak` int(12) NOT NULL,
  `tanggal` date NOT NULL,
  `ket` varchar(100) NOT NULL,
  `id_petugas` int(4) NOT NULL,
  `nm_petugas` varchar(100) NOT NULL,
  `id_penyetor` int(4) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- --------------------------------------------------------

--
-- Table structure for table `data_pbb`
--

CREATE TABLE `data_pbb` (
  `id` int(11) NOT NULL,
  `nop` varchar(20) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `bumi` varchar(5) NOT NULL,
  `bangunan` varchar(5) NOT NULL,
  `pajak` int(12) NOT NULL,
  `alamat_op` varchar(200) NOT NULL,
  `alamat_wp` varchar(200) NOT NULL,
  `ket` varchar(50) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- --------------------------------------------------------

--
-- Table structure for table `data_setor`
--

CREATE TABLE `data_setor` (
  `id` int(11) NOT NULL,
  `kode` varchar(4) NOT NULL,
  `nop` varchar(20) NOT NULL,
  `total_pajak` int(12) NOT NULL,
  `tahun` varchar(4) NOT NULL,
  `rek` varchar(20) NOT NULL,
  `tanggal` date NOT NULL,
  `id_admin` int(4) NOT NULL,
  `status` int(2) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

-- --------------------------------------------------------

--
-- Table structure for table `profil`
--

CREATE TABLE `profil` (
  `id` int(4) NOT NULL,
  `nama_desa` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `kec` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `kab` varchar(75) COLLATE latin1_general_ci NOT NULL,
  `kades` varchar(45) COLLATE latin1_general_ci NOT NULL,
  `kantor` varchar(250) COLLATE latin1_general_ci NOT NULL,
  `email` varchar(150) COLLATE latin1_general_ci NOT NULL,
  `web` varchar(150) COLLATE latin1_general_ci NOT NULL,
  `kode` varchar(4) COLLATE latin1_general_ci NOT NULL,
  `bank` varchar(50) CHARACTER SET latin1 NOT NULL,
  `rek` varchar(18) COLLATE latin1_general_ci NOT NULL,
  `kop1` varchar(200) COLLATE latin1_general_ci NOT NULL,
  `kop2` varchar(200) COLLATE latin1_general_ci NOT NULL,
  `kop3` varchar(200) COLLATE latin1_general_ci NOT NULL,
  `logo` varchar(20) COLLATE latin1_general_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci ROW_FORMAT=DYNAMIC;

--
-- Dumping data for table `profil`
--

INSERT INTO `profil` (`id`, `nama_desa`, `kec`, `kab`, `kades`, `kantor`, `email`, `web`, `kode`, `bank`, `rek`, `kop1`, `kop2`, `kop3`, `logo`) VALUES
(1, 'CONTOH', 'CONTOH', 'CONTOH', 'CONTOH', 'Jl. ..................', 'gudepid@gmail.com', 'https://aplikasidesafree.blogspot.com', '1801', 'BPD LAMPUNG', '123456789', 'PEMERINTAH KABUPATEN CONTOH', 'KECAMATAN CONTOH', 'DESA CONTOH', 'logo.png');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `admin`
--
ALTER TABLE `admin`
  ADD PRIMARY KEY (`id`) USING BTREE,
  ADD KEY `uname` (`uname`) USING BTREE;

--
-- Indexes for table `data_bayar`
--
ALTER TABLE `data_bayar`
  ADD PRIMARY KEY (`id`) USING BTREE,
  ADD UNIQUE KEY `nop` (`nop`) USING BTREE;

--
-- Indexes for table `data_pbb`
--
ALTER TABLE `data_pbb`
  ADD PRIMARY KEY (`id`) USING BTREE,
  ADD UNIQUE KEY `nop_2` (`nop`),
  ADD KEY `nop` (`nop`);

--
-- Indexes for table `data_setor`
--
ALTER TABLE `data_setor`
  ADD PRIMARY KEY (`id`) USING BTREE,
  ADD UNIQUE KEY `nop` (`nop`) USING BTREE;

--
-- Indexes for table `profil`
--
ALTER TABLE `profil`
  ADD PRIMARY KEY (`id`) USING BTREE;

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `admin`
--
ALTER TABLE `admin`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=9;

--
-- AUTO_INCREMENT for table `data_bayar`
--
ALTER TABLE `data_bayar`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `data_pbb`
--
ALTER TABLE `data_pbb`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `data_setor`
--
ALTER TABLE `data_setor`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `profil`
--
ALTER TABLE `profil`
  MODIFY `id` int(4) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=2;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

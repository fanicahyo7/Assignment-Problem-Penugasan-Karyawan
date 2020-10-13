-- phpMyAdmin SQL Dump
-- version 4.8.4
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Aug 29, 2019 at 02:11 AM
-- Server version: 10.1.37-MariaDB
-- PHP Version: 7.3.0

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `cb`
--

-- --------------------------------------------------------

--
-- Table structure for table `karyawan`
--

CREATE TABLE `karyawan` (
  `kd_karyawan` varchar(50) NOT NULL,
  `nama` varchar(50) NOT NULL,
  `jenis_kelamin` varchar(50) NOT NULL,
  `tanggal_lahir` date NOT NULL,
  `alamat` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `karyawan`
--

INSERT INTO `karyawan` (`kd_karyawan`, `nama`, `jenis_kelamin`, `tanggal_lahir`, `alamat`) VALUES
('KYW001', 'Endang', 'Laki-Laki', '1985-04-23', 'Batu'),
('KYW002', 'Sulis', 'Perempuan', '1987-05-13', 'Malang'),
('KYW003', 'Teguh', 'Laki-Laki', '1987-08-20', 'Malang'),
('KYW004', 'Dwi', 'Perempuan', '1986-06-02', 'Batu'),
('KYW005', 'Purwanti', 'Perempuan', '1988-08-19', 'Batu'),
('KYW006', 'Budi', 'Laki-Laki', '1985-01-24', 'Malang'),
('KYW007', 'Titi', 'Perempuan', '1986-12-18', 'Batu');

-- --------------------------------------------------------

--
-- Table structure for table `tb_produksi`
--

CREATE TABLE `tb_produksi` (
  `kd_karyawan` varchar(50) NOT NULL,
  `Tugas_A` int(50) NOT NULL,
  `Tugas_B` int(50) NOT NULL,
  `Tugas_C` int(50) NOT NULL,
  `Tugas_D` int(50) NOT NULL,
  `Tugas_E` int(50) NOT NULL,
  `Tugas_F` int(50) NOT NULL,
  `Tugas_G` int(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_produksi`
--

INSERT INTO `tb_produksi` (`kd_karyawan`, `Tugas_A`, `Tugas_B`, `Tugas_C`, `Tugas_D`, `Tugas_E`, `Tugas_F`, `Tugas_G`) VALUES
('KYW001', 3, 3, 4, 3, 3, 8, 2),
('KYW002', 3, 3, 4, 4, 4, 7, 2),
('KYW003', 2, 4, 4, 3, 4, 9, 3),
('KYW004', 2, 3, 4, 4, 3, 8, 2),
('KYW005', 3, 3, 3, 4, 4, 8, 3),
('KYW006', 3, 3, 4, 3, 4, 8, 3),
('KYW007', 2, 3, 3, 3, 4, 8, 3);

-- --------------------------------------------------------

--
-- Table structure for table `tb_tugasbagi`
--

CREATE TABLE `tb_tugasbagi` (
  `kd_karyawan` varchar(50) NOT NULL,
  `nama_karyawan` varchar(100) NOT NULL,
  `tugas` varchar(25) NOT NULL,
  `bagitugas` varchar(10) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_tugasbagi`
--

INSERT INTO `tb_tugasbagi` (`kd_karyawan`, `nama_karyawan`, `tugas`, `bagitugas`) VALUES
('KYW001', 'Endang', 'B,D,E,G,', 'E'),
('KYW002', 'Sulis', 'B,F,G,', 'F'),
('KYW003', 'Teguh', 'A,D,', 'A'),
('KYW004', 'Dwi', 'A,B,E,G,', 'G'),
('KYW005', 'Purwanti', 'B,C,', 'C'),
('KYW006', 'Budi', 'B,D,', 'B'),
('KYW007', 'Titi', 'A,B,C,D,', 'D');

-- --------------------------------------------------------

--
-- Table structure for table `tb_user`
--

CREATE TABLE `tb_user` (
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_user`
--

INSERT INTO `tb_user` (`Username`, `Password`) VALUES
('admin', 'admin');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `karyawan`
--
ALTER TABLE `karyawan`
  ADD PRIMARY KEY (`kd_karyawan`);

--
-- Indexes for table `tb_produksi`
--
ALTER TABLE `tb_produksi`
  ADD PRIMARY KEY (`kd_karyawan`);

--
-- Indexes for table `tb_tugasbagi`
--
ALTER TABLE `tb_tugasbagi`
  ADD PRIMARY KEY (`kd_karyawan`);

--
-- Indexes for table `tb_user`
--
ALTER TABLE `tb_user`
  ADD PRIMARY KEY (`Username`);

--
-- Constraints for dumped tables
--

--
-- Constraints for table `tb_produksi`
--
ALTER TABLE `tb_produksi`
  ADD CONSTRAINT `tb_produksi_ibfk_1` FOREIGN KEY (`kd_karyawan`) REFERENCES `karyawan` (`kd_karyawan`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

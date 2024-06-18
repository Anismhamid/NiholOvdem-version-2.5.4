-- phpMyAdmin SQL Dump
-- version 5.1.3
-- https://www.phpmyadmin.net/
--
-- Host: localhost
-- Generation Time: Aug 23, 2023 at 10:01 AM
-- Server version: 10.3.38-MariaDB-0ubuntu0.20.04.1
-- PHP Version: 8.0.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `amd_yttd`
--

-- --------------------------------------------------------

--
-- Table structure for table `tloshim`
--

CREATE TABLE `tloshim` (
  `تاريخ الاستلام` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `الى تاريخ` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `من تاريخ` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `الهوية` int(11) NOT NULL,
  `اسم العامل` varchar(30) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Table structure for table `workers`
--

CREATE TABLE `workers` (
  `تعليق` varchar(500) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `الاجر` varchar(50) DEFAULT NULL,
  `عنوان العمل` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `مدير العمل` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `اسم الشركة` varchar(100) DEFAULT NULL,
  `التاريخ` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `الهوية` int(9) DEFAULT NULL,
  `الهاتف` varchar(10) DEFAULT NULL,
  `اسم العامل` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

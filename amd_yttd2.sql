-- phpMyAdmin SQL Dump
-- version 5.1.3
-- https://www.phpmyadmin.net/
--
-- Host: localhost
-- Generation Time: אוקטובר 02, 2023 בזמן 09:59 AM
-- גרסת שרת: 10.3.38-MariaDB-0ubuntu0.20.04.1
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
-- מבנה טבלה עבור טבלה `tloshim`
--

CREATE TABLE `tloshim` (
  `takentade` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `todate` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `fromdate` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `id` int(11) NOT NULL,
  `workername` varchar(30) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- מבנה טבלה עבור טבלה `workers`
--

CREATE TABLE `workers` (
  `note` varchar(500) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `wage` varchar(50) DEFAULT NULL,
  `address` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `managername` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `compname` varchar(100) DEFAULT NULL,
  `date` varchar(50) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
  `id` int(9) DEFAULT NULL,
  `phone` varchar(10) DEFAULT NULL,
  `workername` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

--
-- הוצאת מידע עבור טבלה `workers`
--

INSERT INTO `workers` (`note`, `wage`, `address`, `managername`, `compname`, `date`, `id`, `phone`, `workername`) VALUES
('', '200', '', '', '', '', 2023, '', 'מינוס'),
(),
(),
(),
(),
(),
(),
(),
(),
();
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

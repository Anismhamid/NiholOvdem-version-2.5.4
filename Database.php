<?php
// Database connection
$conn = new mysqli($servername, $username, $password, $dbname, $port);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Receive data from the application and sanitize
$workername = $conn->real_escape_string($_POST['workername']);
$phone = $conn->real_escape_string($_POST['phone']);
$Id = $conn->real_escape_string($_POST['Id']);
$date = $conn->real_escape_string($_POST['date']);
$compname = $conn->real_escape_string($_POST['compname']);
$manager = $conn->real_escape_string($_POST['manager']);
$address = $conn->real_escape_string($_POST['address']);
$wage = $conn->real_escape_string($_POST['wage']);
$note = $conn->real_escape_string($_POST['note']);
$_iid = $conn->real_escape_string($_POST['_iid'])

// SQL query to insert data
$sql = "INSERT INTO workers (`workername`, `phone`, `id`, `date`, `compname`, `manager`, `address`, `wage`, `note`, `_iid`) 
        VALUES ('$workername', '$phone', '$Id', '$date', '$compname', '$manager', '$address', '$wage', '$note', `null`)";

// Execute query and handle the result
if ($conn->query($sql) === TRUE) {
    echo "Data added successfully";
} else {
    echo "Error: " . $conn->error;
}

// Close the database connection
$conn->close();
?>

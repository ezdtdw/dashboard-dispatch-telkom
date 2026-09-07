<?php
header('Content-Type: application/json');
include 'koneksi.php';

$result = $conn->query("SELECT * FROM data_order");
$data = [];

while ($row = $result->fetch_assoc()) {
    $data[] = [
        "TGL" => $row['tanggal'],
        "TEKNISI" => $row['teknisi'],
        "STO" => $row['sto'],
        "NO ORDER" => $row['no_order'],
        "TYPE ORDER" => $row['type_order'],
        "WONUM" => $row['wonum'],
        "NO SERVICE" => $row['no_service'],
        "NAMA" => $row['nama'],
        "CP" => $row['cp'],
        "ALAMAT" => $row['alamat'],
        "TGL CREATE" => $row['tgl_create'],
        "TGL MANJA" => $row['tgl_manja'],
        "ODP" => $row['odp'],
        "PAKET" => $row['paket'],
        "HOMEPASS ID" => $row['homepass_id'],
        "STATUS ORDER" => $row['status_order'],
        "SUB KENDALA" => $row['sub_kendala'],
        "KETERANGAN" => $row['keterangan'],
        "MITRA" => $row['mitra'],
        "SA" => $row['sa']
    ];
}

echo json_encode($data);
$conn->close();
?>
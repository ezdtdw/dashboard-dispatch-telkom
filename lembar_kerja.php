<?php
header('Content-Type: application/json; charset=utf-8');
require_once 'koneksi.php';

if ($conn->connect_error) {
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'Koneksi database gagal.']);
    exit;
}

/*
 * Tabel terpisah supaya upload.php yang melakukan TRUNCATE data_order
 * tidak menghapus data lembar kerja yang sudah diisi.
 */
$createWorksheet = "
CREATE TABLE IF NOT EXISTS lembar_kerja (
    id INT AUTO_INCREMENT PRIMARY KEY,
    wonum VARCHAR(100) NOT NULL UNIQUE,
    sn_ont VARCHAR(150) DEFAULT NULL,
    port_odp VARCHAR(50) DEFAULT NULL,
    koordinat VARCHAR(100) DEFAULT NULL,
    hasil_pekerjaan VARCHAR(50) DEFAULT NULL,
    kendala TEXT DEFAULT NULL,
    keterangan TEXT DEFAULT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
";

$createHistory = "
CREATE TABLE IF NOT EXISTS lembar_kerja_history (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    wonum VARCHAR(100) NOT NULL,
    sn_ont VARCHAR(150) DEFAULT NULL,
    port_odp VARCHAR(50) DEFAULT NULL,
    koordinat VARCHAR(100) DEFAULT NULL,
    hasil_pekerjaan VARCHAR(50) DEFAULT NULL,
    kendala TEXT DEFAULT NULL,
    keterangan TEXT DEFAULT NULL,
    saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_history_wonum (wonum)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
";

if (!$conn->query($createWorksheet) || !$conn->query($createHistory)) {
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'Gagal menyiapkan tabel lembar kerja: ' . $conn->error]);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    $wonum = trim($_GET['wonum'] ?? '');

    if ($wonum === '') {
        echo json_encode(['success' => true, 'data' => []]);
        exit;
    }

    if (isset($_GET['history']) && $_GET['history'] === '1') {
        $stmt = $conn->prepare("
            SELECT sn_ont, port_odp, koordinat, hasil_pekerjaan, kendala, keterangan, saved_at
            FROM lembar_kerja_history
            WHERE wonum = ?
            ORDER BY saved_at DESC
            LIMIT 20
        ");
        $stmt->bind_param('s', $wonum);
        $stmt->execute();
        $historyResult = $stmt->get_result();
        $history = [];
        while ($row = $historyResult->fetch_assoc()) {
            $history[] = $row;
        }
        $stmt->close();

        echo json_encode(['success' => true, 'data' => $history], JSON_UNESCAPED_UNICODE);
        exit;
    }

    $stmt = $conn->prepare("
        SELECT wonum, sn_ont, port_odp, koordinat, hasil_pekerjaan, kendala, keterangan, updated_at
        FROM lembar_kerja
        WHERE wonum = ?
        LIMIT 1
    ");
    $stmt->bind_param('s', $wonum);
    $stmt->execute();

    $result = $stmt->get_result();
    $data = $result->fetch_assoc() ?: [];

    echo json_encode(['success' => true, 'data' => $data], JSON_UNESCAPED_UNICODE);
    $stmt->close();
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!is_array($input)) {
        $input = $_POST;
    }

    $wonum = trim($input['wonum'] ?? '');

    if ($wonum === '') {
        http_response_code(400);
        echo json_encode(['success' => false, 'message' => 'WONUM wajib diisi.']);
        exit;
    }

    $snOnt = trim($input['sn_ont'] ?? '');
    $portOdp = trim($input['port_odp'] ?? '');
    $koordinat = trim($input['koordinat'] ?? '');
    $hasil = trim($input['hasil_pekerjaan'] ?? '');
    $kendala = trim($input['kendala'] ?? '');
    $keterangan = trim($input['keterangan'] ?? '');

    $stmt = $conn->prepare("
        INSERT INTO lembar_kerja
            (wonum, sn_ont, port_odp, koordinat, hasil_pekerjaan, kendala, keterangan)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
            sn_ont = VALUES(sn_ont),
            port_odp = VALUES(port_odp),
            koordinat = VALUES(koordinat),
            hasil_pekerjaan = VALUES(hasil_pekerjaan),
            kendala = VALUES(kendala),
            keterangan = VALUES(keterangan)
    ");

    $stmt->bind_param(
        'sssssss',
        $wonum,
        $snOnt,
        $portOdp,
        $koordinat,
        $hasil,
        $kendala,
        $keterangan
    );

    if (!$stmt->execute()) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => 'Gagal menyimpan: ' . $stmt->error]);
        $stmt->close();
        exit;
    }

    $stmt->close();

    // Simpan snapshot agar perubahan lembar kerja punya riwayat.
    $history = $conn->prepare("
        INSERT INTO lembar_kerja_history
            (wonum, sn_ont, port_odp, koordinat, hasil_pekerjaan, kendala, keterangan)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ");
    $history->bind_param(
        'sssssss',
        $wonum,
        $snOnt,
        $portOdp,
        $koordinat,
        $hasil,
        $kendala,
        $keterangan
    );
    $history->execute();
    $history->close();

    echo json_encode([
        'success' => true,
        'message' => 'Lembar kerja berhasil disimpan.'
    ], JSON_UNESCAPED_UNICODE);
    exit;
}

http_response_code(405);
echo json_encode(['success' => false, 'message' => 'Method tidak didukung.']);
?>

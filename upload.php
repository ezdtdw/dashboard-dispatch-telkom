<?php
include 'koneksi.php';

if (isset($_FILES['file']['tmp_name'])) {
    $filePath = $_FILES['file']['tmp_name'];
    
    // Hapus data lama agar diganti data baru
    $conn->query("TRUNCATE TABLE data_order");
    
    // Fungsi merapikan tanggal (Otomatis membuang jam jika ada)
    function parseTanggal($val) {
        if(empty($val) || trim($val) == '-') return date('Y-m-d');
        
        // Buang bagian jam (pisahkan berdasarkan spasi, ambil yang pertama saja)
        // Contoh: "9/1/2026 14:30:00" -> akan diambil "9/1/2026" saja
        $val = explode(' ', trim($val))[0];

        // Cek jika formatnya ada garis miring (M/D/YYYY)
        if (strpos($val, '/') !== false) {
            $parts = explode('/', $val);
            // Susun ulang jadi YYYY-MM-DD
            if (count($parts) == 3) {
                $tahun = strlen($parts[2]) == 2 ? '20'.$parts[2] : $parts[2];
                $bulan = str_pad($parts[0], 2, "0", STR_PAD_LEFT);
                $hari = str_pad($parts[1], 2, "0", STR_PAD_LEFT);
                return "$tahun-$bulan-$hari";
            }
        }
        
        // Kalau formatnya bukan M/D/YYYY, serahkan ke sistem PHP biasa
        $time = strtotime(str_replace('/', '-', $val));
        return $time ? date('Y-m-d', $time) : date('Y-m-d');
    }

    // Siapkan query dengan 20 kolom (sebelumnya cuma 13)
    $stmt = $conn->prepare("INSERT INTO data_order (tanggal, teknisi, sto, no_order, type_order, wonum, no_service, nama, cp, alamat, tgl_create, tgl_manja, odp, paket, homepass_id, status_order, sub_kendala, keterangan, mitra, sa) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

    // Deteksi pemisah CSV secara otomatis (Koma atau Titik Koma)
    $file = fopen($filePath, "r");
    $firstLine = fgets($file);
    $separator = strpos($firstLine, ';') !== false ? ';' : ',';
    rewind($file);

    $row = 0;
    while (($data = fgetcsv($file, 10000, $separator)) !== FALSE) {
        $row++;
        if ($row == 1) continue; // Lewati baris 1 (Judul/Header)
        if (empty($data[4]) && empty($data[1])) continue; // Lewati baris kosong

        // Mapping Kolom CSV Baru
        $tanggal     = parseTanggal($data[0] ?? '');                   
        $teknisi     = !empty($data[1]) ? trim($data[1]) : '-';        
        $sto         = !empty($data[3]) ? trim($data[3]) : '-';        
        $no_order    = !empty($data[4]) ? trim($data[4]) : '-';        
        $type_order  = !empty($data[5]) ? trim($data[5]) : '-';        // BARU
        $wonum       = !empty($data[6]) ? trim($data[6]) : '-';        
        $no_service  = !empty($data[7]) ? trim($data[7]) : '-';        // BARU
        $nama        = !empty($data[8]) ? trim($data[8]) : '-';        
        $cp          = !empty($data[9]) ? trim($data[9]) : '-';        // BARU
        $alamat      = !empty($data[10]) ? trim($data[10]) : '-';      // BARU
        $tgl_create  = parseTanggal($data[11] ?? '');                  
        $tgl_manja   = !empty($data[12]) ? trim($data[12]) : '-';      // BARU
        $odp         = !empty($data[13]) ? trim($data[13]) : '-';      
        $paket       = !empty($data[14]) ? trim($data[14]) : '-';      // BARU
        $homepass_id = !empty($data[17]) ? trim($data[17]) : '-';      // BARU
        $status      = !empty($data[18]) ? trim($data[18]) : 'OGP';    
        $sub_kendala = !empty($data[19]) ? trim($data[19]) : '-';      
        $keterangan  = !empty($data[20]) ? trim($data[20]) : '-';      
        $mitra       = !empty($data[21]) ? trim($data[21]) : '-';      
        $sa          = !empty($data[22]) ? trim($data[22]) : '-';      

        // Eksekusi data ke DB (ada 20 huruf 's')
        $stmt->bind_param("ssssssssssssssssssss", $tanggal, $teknisi, $sto, $no_order, $type_order, $wonum, $no_service, $nama, $cp, $alamat, $tgl_create, $tgl_manja, $odp, $paket, $homepass_id, $status, $sub_kendala, $keterangan, $mitra, $sa);
        $stmt->execute();
    }
    
    fclose($file);
    echo "Berhasil";
}
?>
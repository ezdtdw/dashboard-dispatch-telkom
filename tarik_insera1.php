<?php
include 'koneksi.php';

$url = 'https://wfm.telkom.co.id/jw/web/userview/new_wfm/v/_/inboxworkorder2?d-28330-ps=100';

$ch = curl_init();
curl_setopt($ch, CURLOPT_URL,$url);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_ENCODING, '');
curl_setopt($ch, CURLOPT_MAXREDIRS, 10);
curl_setopt($ch, CURLOPT_TIMEOUT, 0);
curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'GET');

// PASTIKAN COOKIE INI MASIH AKTIF ATAU UPDATE DENGAN YANG BARU DARI MAS MARGO
$headers = [
  'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
  'Accept-Language: en-US,en;q=0.9,id;q=0.8',
  'Connection: keep-alive',
  'Cookie: JSESSIONID=Py9C1vrdRIJrXGxBuUk2u19keXoYmXCLzIUrTzCK.bima-7d5666fcd6-862vk; _hjSessionUser_6634322=eyJpZCI6IjQ2MjgzYWEyLTNiMDYtNTcyOS1hZGFhLTVmMGIxYmEwY2UzMCIsImNyZWF0ZWQiOjE3ODYzNDAxNTc1NzIsImV4aXN0aW5nIjp0cnVlfQ==; 2583a12ba4a45f8c3321a228b61ba02c=f5196855c5de3ab77ef5d1890cb504bd',
  'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36'
];
curl_setopt($ch, CURLOPT_HTTPHEADER,$headers);

$html_response = curl_exec($ch);
$err = curl_error($ch);
curl_close($ch);

if ($err) {
    die("Error narik Insera: " . $err);
}

libxml_use_internal_errors(true);
$dom = new DOMDocument();
$dom->loadHTML($html_response);
libxml_clear_errors();

$xpath = new DOMXPath($dom);
$rows =$xpath->query('//table[@id="inbox_workorder"]//tbody//tr');

$total_insert = 0;
$total_update = 0;

foreach ($rows as $row) {$cols = $xpath->query('td',$row);
    
    if ($cols->length > 15) { 
        // Mengambil data dari Insera sesuai index
        $tgl_create   = trim($cols->item(1)->nodeValue ?? ''); // Date Created
        $tanggal      = trim($cols->item(2)->nodeValue ?? ''); // Status Date (Jadi tanggal update)
        $wonum        = trim($cols->item(4)->nodeValue ?? ''); // Workorder
        $no_order     = trim($cols->item(5)->nodeValue ?? ''); // SC Order No
        $type_order   = trim($cols->item(8)->nodeValue ?? ''); // CRM Order Type
        $mitra        = trim($cols->item(9)->nodeValue ?? ''); // Owner Group
        $status_order = trim($cols->item(10)->nodeValue ?? ''); // Status
        $nama         = trim($cols->item(12)->nodeValue ?? ''); // Customer Name
        $sto          = trim($cols->item(13)->nodeValue ?? ''); // Workzone
        $tgl_manja    = trim($cols->item(16)->nodeValue ?? ''); // Sched Start
        $no_service   = trim($cols->item(17)->nodeValue ?? ''); // Service Num
        $cp           = trim($cols->item(18)->nodeValue ?? ''); // Contact Number
        
        $wonum = preg_replace('/\s+/', '',$wonum);

        if (empty($wonum)) continue;

        // if (strpos(strtoupper($mitra), 'SAMARINDA') === false) {
        //     continue;
        // }

        // Cek data di MySQL berdasarkan nama kolom milikmu
        $cek = $conn->query("SELECT wonum FROM data_order WHERE wonum = '$wonum'");
        
        if ($cek && $cek->num_rows > 0) {
            
            // Kolom status_order pakai underscore
            $sql_update = "UPDATE data_order SET status_order = '$status_order', tanggal = '$tanggal', sto = '$sto', mitra = '$mitra' WHERE wonum = '$wonum'";
            $conn->query($sql_update);
            $total_update++;
            
        } else {
            //
            // Semua kolom yang tadinya spasi diganti underscore (no_order, type_order, no_service, dll)
            $sql_insert = "INSERT INTO data_order (tanggal, teknisi, sto, no_order, type_order, wonum, no_service, nama, cp, alamat, tgl_create, tgl_manja, odp, paket, homepass_id, status_order, sub_kendala, keterangan, mitra, sa) VALUES ('$tanggal', '', '$sto', '$no_order', '$type_order', '$wonum', '$no_service', '$nama', '$cp', '', '$tgl_create', '$tgl_manja', '', '', '', '$status_order', '', '', '$mitra', '')";
            $conn->query($sql_insert);
            $total_insert++;
            // 
        }
    }
}

echo "Berhasil! Data baru ditambahkan: $total_insert \vert{} Data lama diupdate:$total_update";
?>
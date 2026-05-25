<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ThuocController extends Controller
{
    public function index()
    {
        return view('add_ma_thuoc_dv');
    }

    private function escapeSqlString($value)
    {
        if ($value === null) return "NULL";
        $escaped = str_replace("'", "''", $value);
        return "N'" . $escaped . "'";
    }

    public function process(Request $request)
    {
        $request->validate(['excel_file' => 'required|mimes:xlsx,xls,csv']);
        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray();

        $queries = [];
        for ($i = 2; $i < count($rows); $i++) {
            $row = $rows[$i];
            if (empty($row[1])) continue;

            $maThuoc = $row[1];
            $tenThuoc = $this->escapeSqlString($row[4]);
            $donViTinh = $this->escapeSqlString($row[5]);
            $hamLuong = $this->escapeSqlString($row[6]);
            $cachDung = $this->escapeSqlString($row[7]);
            $hangSanXuat = $this->escapeSqlString($row[8]);
            $nuocSanXuat = $this->escapeSqlString($row[9]);
            $quyCach = $this->escapeSqlString($row[10]);
            $hoatChat = $this->escapeSqlString($row[12]);
            $giaMua = $row[14] ?? 0;
            $giaBan = $row[15] ?? 0;
            $coBH = ($row[16] === 'x' || $row[16] === 'X') ? 'true' : 'false';

            $queries[] = "INSERT INTO \"Thuoc\" (\"MaThuoc\", \"TenThuoc\", \"DonViTinh\", \"HamLuong\", \"CachDung\", \"HangSanXuat\", \"NuocSanXuat\", \"QuyCach\", \"HoatChat\", \"GiaMua\", \"GiaBan\", \"CoBH\", \"Xoa\", \"NhaThuocNgoai\") VALUES ('$maThuoc', $tenThuoc, $donViTinh, $hamLuong, $cachDung, $hangSanXuat, $nuocSanXuat, $quyCach, $hoatChat, $giaMua, $giaBan, $coBH, false, false);";
        }

        return response()->json([
            'success' => true,
            'init_queries' => "CREATE TEMP TABLE tbTam (MaThuoc varchar(255));",
            'queries' => implode("\n", $queries),
            'warning_queries' => "INSERT INTO \"Thuoc_CanhBaoTon\" (\"IDThuoc\", \"SoLuong\", \"MaKho\") SELECT \"IDThuoc\", 0, 2 FROM \"Thuoc\" WHERE \"CoBH\"=false AND \"Xoa\"=false AND \"IDThuoc\" NOT IN (SELECT \"IDThuoc\" FROM \"Thuoc_CanhBaoTon\" WHERE \"MaKho\"=2);"
        ]);
    }

    public function processKiGui(Request $request)
    {
        $request->validate(['excel_file' => 'required|mimes:xlsx,xls,csv']);
        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray();

        $queries = [];
        $maThuocList = [];
        for ($i = 2; $i < count($rows); $i++) {
            $row = $rows[$i];
            if (empty($row[1])) continue;

            $maThuoc = $row[1];
            $maThuocList[] = $maThuoc;
            $tenThuoc = $this->escapeSqlString($row[4]);
            $donViTinh = $this->escapeSqlString($row[5]);
            $hamLuong = $this->escapeSqlString($row[6]);
            $cachDung = $this->escapeSqlString($row[7]);
            $hangSanXuat = $this->escapeSqlString($row[8]);
            $nuocSanXuat = $this->escapeSqlString($row[9]);
            $quyCach = $this->escapeSqlString($row[10]);
            $hoatChat = $this->escapeSqlString($row[12]);
            $giaMua = $row[14] ?? 0;
            $giaBan = $row[15] ?? 0;
            $coBH = ($row[16] === 'x' || $row[16] === 'X') ? 'true' : 'false';

            $queries[] = "INSERT INTO \"Thuoc\" (\"MaThuoc\", \"TenThuoc\", \"DonViTinh\", \"HamLuong\", \"CachDung\", \"HangSanXuat\", \"NuocSanXuat\", \"QuyCach\", \"HoatChat\", \"GiaMua\", \"GiaBan\", \"CoBH\", \"Xoa\", \"NhaThuocNgoai\") VALUES ('$maThuoc', $tenThuoc, $donViTinh, $hamLuong, $cachDung, $hangSanXuat, $nuocSanXuat, $quyCach, $hoatChat, $giaMua, $giaBan, $coBH, false, false);";
        }

        $maThuocIn = "'" . implode("','", $maThuocList) . "'";

        return response()->json([
            'success' => true,
            'init_queries' => "CREATE TEMP TABLE tbTam (MaThuoc varchar(255));",
            'queries' => implode("\n", $queries),
            'warning_queries' => "INSERT INTO \"Thuoc_CanhBaoTon\" (\"IDThuoc\", \"SoLuong\", \"MaKho\") SELECT \"IDThuoc\", 0, 1 FROM \"Thuoc\" WHERE \"CoBH\"=false AND \"Xoa\"=false AND \"IDThuoc\" NOT IN (SELECT \"IDThuoc\" FROM \"Thuoc_CanhBaoTon\" WHERE \"MaKho\"=1);",
            'update_kigui_queries' => "UPDATE \"Thuoc\" SET \"NhaThuocNgoai\"=true, \"IDDVCTN\"=3 WHERE \"MaThuoc\" IN ($maThuocIn);",
            'check_queries' => "SELECT \"NhaThuocNgoai\", \"IDDVCTN\", * FROM \"Thuoc\" WHERE \"MaThuoc\" IN ($maThuocIn);"
        ]);
    }

    public function processVTYTDV(Request $request)
    {
        $request->validate(['excel_file' => 'required|mimes:xlsx,xls,csv']);
        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray();

        $queries = [];
        for ($i = 2; $i < count($rows); $i++) {
            $row = $rows[$i];
            if (empty($row[1])) continue;

            $maVTYT = $row[1];
            $tenVTYT = $this->escapeSqlString($row[4]);
            $donViTinh = $this->escapeSqlString($row[5]);
            $quyCach = $this->escapeSqlString($row[6]);
            $hangSanXuat = $this->escapeSqlString($row[7]);
            $nuocSanXuat = $this->escapeSqlString($row[9]);

            $queries[] = "INSERT INTO \"VatTuYTe\" (\"MaVTYT\", \"TenVTYT\", \"DonViTinh\", \"QuyCach\", \"HangSanXuat\", \"NuocSanXuat\", \"Xoa\", \"NhaThuocNgoai\", \"CoBH\") VALUES ('$maVTYT', $tenVTYT, $donViTinh, $quyCach, $hangSanXuat, $nuocSanXuat, false, false, false);";
        }

        return response()->json([
            'success' => true,
            'init_queries' => "DROP TABLE IF EXISTS tbTam;",
            'queries' => implode("\n", $queries),
            'warning_queries' => "INSERT INTO \"VTYT_CanhBaoTon\" (\"IDVTYT\", \"SoLuong\", \"MaKho\") SELECT \"IDVTYT\", 0, 2 FROM \"VatTuYTe\" WHERE \"CoBH\"=false AND \"Xoa\"=false AND \"IDVTYT\" NOT IN (SELECT \"IDVTYT\" FROM \"VTYT_CanhBaoTon\" WHERE \"MaKho\"=2);"
        ]);
    }

    public function processVTYTKiGui(Request $request)
    {
        $request->validate(['excel_file' => 'required|mimes:xlsx,xls,csv']);
        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray();

        $queries = [];
        $maVTYTList = [];
        for ($i = 2; $i < count($rows); $i++) {
            $row = $rows[$i];
            if (empty($row[1])) continue;

            $maVTYT = $row[1];
            $maVTYTList[] = $maVTYT;
            $tenVTYT = $this->escapeSqlString($row[4]);
            $donViTinh = $this->escapeSqlString($row[5]);
            $quyCach = $this->escapeSqlString($row[6]);
            $hangSanXuat = $this->escapeSqlString($row[7]);
            $nuocSanXuat = $this->escapeSqlString($row[9]);

            $queries[] = "INSERT INTO \"VatTuYTe\" (\"MaVTYT\", \"TenVTYT\", \"DonViTinh\", \"QuyCach\", \"HangSanXuat\", \"NuocSanXuat\", \"Xoa\", \"NhaThuocNgoai\", \"CoBH\") VALUES ('$maVTYT', $tenVTYT, $donViTinh, $quyCach, $hangSanXuat, $nuocSanXuat, false, false, false);";
        }

        $maVTYTIn = "'" . implode("','", $maVTYTList) . "'";

        return response()->json([
            'success' => true,
            'init_queries' => "DROP TABLE IF EXISTS tbTam;",
            'queries' => implode("\n", $queries),
            'warning_queries' => "INSERT INTO \"VTYT_CanhBaoTon\" (\"IDVTYT\", \"SoLuong\", \"MaKho\") SELECT \"IDVTYT\", 0, 1 FROM \"VatTuYTe\" WHERE \"CoBH\"=false AND \"Xoa\"=false AND \"IDVTYT\" NOT IN (SELECT \"IDVTYT\" FROM \"VTYT_CanhBaoTon\" WHERE \"MaKho\"=1);",
            'update_kigui_queries' => "UPDATE \"VatTuYTe\" SET \"NhaThuocNgoai\"=true, \"IDDVCTN\"=3 WHERE \"MaVTYT\" IN ($maVTYTIn);",
            'check_queries' => "SELECT * FROM \"VatTuYTe\" WHERE \"MaVTYT\" IN ($maVTYTIn);"
        ]);
    }
}

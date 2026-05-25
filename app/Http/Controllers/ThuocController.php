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

    // --- THUỐC DỊCH VỤ ---
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
            if (empty($row[4])) continue; // Kiểm tra Mã Thuốc (Cột E)

            $maThuoc = $row[4];               // Cột E
            $tenThuoc = $this->escapeSqlString($row[5]); // Cột F
            $tenKhoaHoc = $this->escapeSqlString($row[6]); // Cột G
            $hoatChat = $this->escapeSqlString($row[7]);   // Cột H
            $nhaSX = $this->escapeSqlString($row[8]);     // Cột I
            $duongDung = $this->escapeSqlString($row[9]); // Cột J
            $dvt = $this->escapeSqlString($row[10]);      // Cột K
            $quiCach = $this->escapeSqlString($row[12]);  // Cột M
            $maNT = $row[1] ?? 'NULL';        // Cột B
            $nhomThuoc = $row[14] ?? 'NULL';  // Cột O
            $dangDung = $row[15] ?? 'NULL';   // Cột P
            $thuocDonChat = (isset($row[16]) && (strtolower($row[16]) === 'x')) ? 'true' : 'false'; // Cột Q

            $queries[] = "INSERT INTO \"Thuoc\"(\"MaThuoc\",\"TenThuoc\",\"TenKhoaHoc\", \"HoatChat\",\"NhaSX\",\"DuongDung\",\"DVT\", \"QuiCach\",\"MaNT\",\"NhomThuoc\",\"DangDung\",\"ThuocDonChat\") VALUES ('$maThuoc', $tenThuoc, $tenKhoaHoc, $hoatChat, $nhaSX, $duongDung, $dvt, $quiCach, $maNT, $nhomThuoc, $dangDung, $thuocDonChat);";
        }

        $initSql = "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDThuoc\",\"MaKhoa\", cast(\"IDThuoc\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"Thuoc\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongThuocTon\"(\"IDThuoc\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDThuoc\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongThuocTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDThuoc\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;";

        $warningSql = "insert into \"Thuoc_CanhBaoTon\" ( \"IDThuoc\", \"SoLuong\" , \"MaKho\")\nselect \"IDThuoc\",0,2\nfrom \"Thuoc\" where \"ThuocBH\"=false AND \"Xoa\"=false and \"IDThuoc\" NOT IN (select \"IDThuoc\"from \"Thuoc_CanhBaoTon\" where \"MaKho\"=2)";

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => $initSql,
            'warning_queries' => $warningSql
        ]);
    }

    // --- THUỐC KÍ GỬI ---
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
            if (empty($row[4])) continue; // Kiểm tra Mã Thuốc (Cột E)

            $maThuoc = $row[4];               // Cột E
            $maThuocList[] = $maThuoc;
            $tenThuoc = $this->escapeSqlString($row[5]); // Cột F
            $tenKhoaHoc = $this->escapeSqlString($row[6]); // Cột G
            $hoatChat = $this->escapeSqlString($row[7]);   // Cột H
            $nhaSX = $this->escapeSqlString($row[8]);     // Cột I
            $duongDung = $this->escapeSqlString($row[9]); // Cột J
            $dvt = $this->escapeSqlString($row[10]);      // Cột K
            $quiCach = $this->escapeSqlString($row[12]);  // Cột M
            $maNT = $row[1] ?? 'NULL';        // Cột B
            $nhomThuoc = $row[14] ?? 'NULL';  // Cột O
            $dangDung = $row[15] ?? 'NULL';   // Cột P
            $thuocDonChat = (isset($row[16]) && (strtolower($row[16]) === 'x')) ? 'true' : 'false'; // Cột Q

            $queries[] = "INSERT INTO \"Thuoc\"(\"MaThuoc\",\"TenThuoc\",\"TenKhoaHoc\", \"HoatChat\",\"NhaSX\",\"DuongDung\",\"DVT\", \"QuiCach\",\"MaNT\",\"NhomThuoc\",\"DangDung\",\"ThuocDonChat\") VALUES ('$maThuoc', $tenThuoc, $tenKhoaHoc, $hoatChat, $nhaSX, $duongDung, $dvt, $quiCach, $maNT, $nhomThuoc, $dangDung, $thuocDonChat);";
        }

        $maThuocIn = "'" . implode("','", $maThuocList) . "'";
        $initSql = "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDThuoc\",\"MaKhoa\", cast(\"IDThuoc\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"Thuoc\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongThuocTon\"(\"IDThuoc\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDThuoc\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongThuocTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDThuoc\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;";
        $warningSql = "insert into \"Thuoc_CanhBaoTon\" ( \"IDThuoc\", \"SoLuong\" , \"MaKho\")\nselect \"IDThuoc\",0,1\nfrom \"Thuoc\" where \"ThuocBH\"=false AND \"Xoa\"=false and \"IDThuoc\" NOT IN (select \"IDThuoc\"from \"Thuoc_CanhBaoTon\" where \"MaKho\"=1)";

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => $initSql,
            'warning_queries' => $warningSql,
            'update_kigui_queries' => "update \"Thuoc\" set \"NhaThuocNgoai\"=true,\"IDDVCTN\"=3 WHERE \"MaThuoc\" in ($maThuocIn);",
            'check_queries' => "select \"NhaThuocNgoai\",\"IDDVCTN\", * from \"Thuoc\" WHERE \"MaThuoc\" in ($maThuocIn);"
        ]);
    }

    // --- VTYT DỊCH VỤ ---
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
            if (empty($row[4])) continue; // Kiểm tra Mã VTYT (Cột E)
            
            $maNhomVTYT = $row[1] ?? 'NULL'; // Cột B
            $maVTYT = $row[4];               // Cột E
            $tenVTYT = $this->escapeSqlString($row[6]); // Cột G
            $nhaSX = $this->escapeSqlString($row[7]);   // Cột H
            $dvt = $this->escapeSqlString($row[8]);     // Cột I
            $quiCach = $this->escapeSqlString($row[10]); // Cột K

            $queries[] = "INSERT INTO \"VatTuYTe\"(\"MaVTYT\",\"TenVTYT\",\"NhaSX\",\"DVT\",\"QuiCach\",\"SoLuongCon\",\"CoBH\",\"MaNhomVTYT\") VALUES('$maVTYT',$tenVTYT,$nhaSX,$dvt,$quiCach,0,false,$maNhomVTYT);";
        }

        $initSql = "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDVTYT\",\"MaKhoa\", cast(\"IDVTYT\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"VatTuYTe\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongVTYTTon\"(\"IDVTYT\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDVTYT\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongVTYTTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDVTYT\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;";

        $warningSql = "insert into \"VTYT_CanhBaoTon\" ( \"IDVTYT\", \"SoLuong\" , \"MaKho\")\nselect \"IDVTYT\",0,2\nfrom \"VatTuYTe\" where \"CoBH\"=false AND \"Xoa\"=false and \"IDVTYT\" NOT IN (select \"IDVTYT\"from \"VTYT_CanhBaoTon\" where \"MaKho\"=2)";

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => $initSql,
            'warning_queries' => $warningSql
        ]);
    }

    // --- VTYT KÍ GỬI ---
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
            if (!isset($row[4]) || empty($row[4])) continue; // Kiểm tra Mã VTYT (Cột E)
            
            $maNhomVTYT = $row[1] ?? 'NULL'; // Cột B
            $maVTYT = $row[4];               // Cột E
            $maVTYTList[] = $maVTYT;
            
            $tenVTYT = isset($row[6]) ? $this->escapeSqlString($row[6]) : "N''"; // Cột G
            $nhaSX = isset($row[7]) ? $this->escapeSqlString($row[7]) : "N''";   // Cột H
            $dvt = isset($row[8]) ? $this->escapeSqlString($row[8]) : "N''";     // Cột I
            $quiCach = isset($row[10]) ? $this->escapeSqlString($row[10]) : "N''"; // Cột K

            $queries[] = "INSERT INTO \"VatTuYTe\"(\"MaVTYT\",\"TenVTYT\",\"NhaSX\",\"DVT\",\"QuiCach\",\"SoLuongCon\",\"CoBH\",\"MaNhomVTYT\") VALUES('$maVTYT',$tenVTYT,$nhaSX,$dvt,$quiCach,0,false,$maNhomVTYT);";
        }

        $maVTYTIn = "'" . implode("','", $maVTYTList) . "'";
        $initSql = "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDVTYT\",\"MaKhoa\", cast(\"IDVTYT\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"VatTuYTe\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongVTYTTon\"(\"IDVTYT\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDVTYT\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongVTYTTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDVTYT\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;";
        $warningSql = "insert into \"VTYT_CanhBaoTon\" ( \"IDVTYT\", \"SoLuong\" , \"MaKho\")\nselect \"IDVTYT\",0,1\nfrom \"VatTuYTe\" where \"CoBH\"=false AND \"Xoa\"=false and \"IDVTYT\" NOT IN (select \"IDVTYT\"from \"VTYT_CanhBaoTon\" where \"MaKho\"=1)";

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => $initSql,
            'warning_queries' => $warningSql,
            'update_kigui_queries' => "update \"VatTuYTe\" set \"NhaThuocNgoai\"=true,\"IDDVCTN\"=3 where \"MaVTYT\" in ($maVTYTIn);",
            'check_queries' => "select * from \"VatTuYTe\" where \"MaVTYT\" in ($maVTYTIn);"
        ]);
    }
}

<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class ThuocController extends Controller
{
    public function index()
    {
        return view('add_ma_thuoc_dv');
    }

    public function process(Request $request)
    {
        // ... (giữ nguyên logic cũ cho Dịch Vụ)
        $request->validate([
            'excel_file' => 'required|mimes:xlsx,xls,csv'
        ]);

        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        
        $queries = [];
        
        for ($row = 3; $row <= $highestRow; $row++) {
            $maNT = $worksheet->getCell('B' . $row)->getValue();
            $maThuoc = $worksheet->getCell('E' . $row)->getValue();
            $tenThuoc = $worksheet->getCell('F' . $row)->getValue();
            $tenKhoaHoc = $worksheet->getCell('G' . $row)->getValue();
            $hoatChat = $worksheet->getCell('H' . $row)->getValue();
            $nhaSX = $worksheet->getCell('I' . $row)->getValue();
            $duongDung = $worksheet->getCell('J' . $row)->getValue();
            $dvt = $worksheet->getCell('K' . $row)->getValue();
            $quiCach = $worksheet->getCell('M' . $row)->getValue();
            $nhomThuoc = $worksheet->getCell('O' . $row)->getValue();
            $dangDung = $worksheet->getCell('P' . $row)->getValue();
            $donChat = strtolower(trim($worksheet->getCell('Q' . $row)->getValue())) === 'x' ? 'true' : 'false';

            if (empty($maThuoc)) continue;

            $tenThuocEsc = str_replace("'", "''", $tenThuoc);
            $tenKhoaHocEsc = str_replace("'", "''", $tenKhoaHoc);
            $hoatChatEsc = str_replace("'", "''", $hoatChat);
            $nhaSXEsc = str_replace("'", "''", $nhaSX);
            $duongDungEsc = str_replace("'", "''", $duongDung);
            $dvtEsc = str_replace("'", "''", $dvt);
            $quiCachEsc = str_replace("'", "''", $quiCach);

            $query = "INSERT INTO \"Thuoc\"(\"MaThuoc\",\"TenThuoc\",\"TenKhoaHoc\", \"HoatChat\",\"NhaSX\",\"DuongDung\",\"DVT\", \"QuiCach\",\"MaNT\",\"NhomThuoc\",\"DangDung\",\"ThuocDonChat\")\n";
            $query .= " VALUES ('$maThuoc', N'$tenThuocEsc', N'$tenKhoaHocEsc', N'$hoatChatEsc', N'$nhaSXEsc', N'$duongDungEsc', N'$dvtEsc', N'$quiCachEsc', $maNT, $nhomThuoc, $dangDung, $donChat);";
            
            $queries[] = $query;
        }

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDThuoc\",\"MaKhoa\", cast(\"IDThuoc\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"Thuoc\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongThuocTon\"(\"IDThuoc\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDThuoc\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongThuocTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDThuoc\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;",
            'warning_queries' => "insert into \"Thuoc_CanhBaoTon\" ( \"IDThuoc\",  \"SoLuong\" ,  \"MaKho\")\nselect  \"IDThuoc\",0,2\nfrom \"Thuoc\"  where \"ThuocBH\"=false AND \"Xoa\"=false and \"IDThuoc\" NOT IN (select \"IDThuoc\"from \"Thuoc_CanhBaoTon\" where \"MaKho\"=2)"
        ]);
    }

    public function processKiGui(Request $request)
    {
        // ... (existing processKiGui logic)
        $request->validate([
            'excel_file' => 'required|mimes:xlsx,xls,csv'
        ]);

        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        
        $queries = [];
        $maThuocList = [];
        
        for ($row = 3; $row <= $highestRow; $row++) {
            $maNT = $worksheet->getCell('B' . $row)->getValue();
            $maThuoc = $worksheet->getCell('E' . $row)->getValue();
            $tenThuoc = $worksheet->getCell('F' . $row)->getValue();
            $tenKhoaHoc = $worksheet->getCell('G' . $row)->getValue();
            $hoatChat = $worksheet->getCell('H' . $row)->getValue();
            $nhaSX = $worksheet->getCell('I' . $row)->getValue();
            $duongDung = $worksheet->getCell('J' . $row)->getValue();
            $dvt = $worksheet->getCell('K' . $row)->getValue();
            $quiCach = $worksheet->getCell('M' . $row)->getValue();
            $nhomThuoc = $worksheet->getCell('O' . $row)->getValue();
            $dangDung = $worksheet->getCell('P' . $row)->getValue();
            $donChat = strtolower(trim($worksheet->getCell('Q' . $row)->getValue())) === 'x' ? 'true' : 'false';

            if (empty($maThuoc)) continue;

            $maThuocList[] = $maThuoc;

            $tenThuocEsc = str_replace("'", "''", $tenThuoc);
            $tenKhoaHocEsc = str_replace("'", "''", $tenKhoaHoc);
            $hoatChatEsc = str_replace("'", "''", $hoatChat);
            $nhaSXEsc = str_replace("'", "''", $nhaSX);
            $duongDungEsc = str_replace("'", "''", $duongDung);
            $dvtEsc = str_replace("'", "''", $dvt);
            $quiCachEsc = str_replace("'", "''", $quiCach);

            // Cấu trúc INSERT tương tự nhưng có thể tùy chỉnh nếu cần trong tương lai
            $query = "INSERT INTO \"Thuoc\"(\"MaThuoc\",\"TenThuoc\",\"TenKhoaHoc\", \"HoatChat\",\"NhaSX\",\"DuongDung\",\"DVT\", \"QuiCach\",\"MaNT\",\"NhomThuoc\",\"DangDung\",\"ThuocDonChat\")\n";
            $query .= " VALUES ('$maThuoc', N'$tenThuocEsc', N'$tenKhoaHocEsc', N'$hoatChatEsc', N'$nhaSXEsc', N'$duongDungEsc', N'$dvtEsc', N'$quiCachEsc', $maNT, $nhomThuoc, $dangDung, $donChat);";
            
            $queries[] = $query;
        }

        $maThuocList = array_map(function($m) { return "'$m'"; }, array_filter($maThuocList));
        $maThuocIn = implode(',', $maThuocList);

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDThuoc\",\"MaKhoa\", cast(\"IDThuoc\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"Thuoc\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongThuocTon\"(\"IDThuoc\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDThuoc\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongThuocTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDThuoc\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;",
            'warning_queries' => "insert into \"Thuoc_CanhBaoTon\" ( \"IDThuoc\",  \"SoLuong\" ,  \"MaKho\")\nselect  \"IDThuoc\",0,1\nfrom \"Thuoc\"  where \"ThuocBH\"=false AND \"Xoa\"=false and \"IDThuoc\" NOT IN (select \"IDThuoc\"from \"Thuoc_CanhBaoTon\" where \"MaKho\"=1)",
            'update_kigui_queries' => "update \"Thuoc\" set \"NhaThuocNgoai\"=true,\"IDDVCTN\"=3\nWHERE \"MaThuoc\" in ($maThuocIn);",
            'check_queries' => "select \"NhaThuocNgoai\",\"IDDVCTN\", * from \"Thuoc\"\nWHERE \"MaThuoc\" in ($maThuocIn);"
        ]);
    }

    public function processVTYTDV(Request $request)
    {
        $request->validate([
            'excel_file' => 'required|mimes:xlsx,xls,csv'
        ]);

        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        
        $queries = [];
        
        for ($row = 3; $row <= $highestRow; $row++) {
            $maNhomVTYT = $worksheet->getCell('B' . $row)->getValue();
            $maVTYT = $worksheet->getCell('E' . $row)->getValue();
            $tenVTYT = $worksheet->getCell('F' . $row)->getValue();
            $nhaSX = $worksheet->getCell('G' . $row)->getValue();
            $dvt = $worksheet->getCell('H' . $row)->getValue();
            $quiCach = $worksheet->getCell('J' . $row)->getValue();

            if (empty($maVTYT)) continue;

            $tenVTYTEsc = str_replace("'", "''", $tenVTYT);
            $nhaSXEsc = str_replace("'", "''", $nhaSX);
            $dvtEsc = str_replace("'", "''", $dvt);
            $quiCachEsc = str_replace("'", "''", $quiCach);

            $query = "INSERT INTO \"VatTuYTe\"(\"MaVTYT\",\"TenVTYT\",\"NhaSX\",\"DVT\",\"QuiCach\",\"SoLuongCon\",\"CoBH\",\"MaNhomVTYT\")\n";
            $query .= "VALUES('$maVTYT',N'$tenVTYTEsc',N'$nhaSXEsc',N'$dvtEsc',N'$quiCachEsc',0,false,$maNhomVTYT);";
            
            $queries[] = $query;
        }

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDVTYT\",\"MaKhoa\", cast(\"IDVTYT\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"VatTuYTe\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongVTYTTon\"(\"IDVTYT\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDVTYT\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongVTYTTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDVTYT\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;",
            'warning_queries' => "insert into \"VTYT_CanhBaoTon\" ( \"IDVTYT\",  \"SoLuong\" ,  \"MaKho\")\nselect   \"IDVTYT\",0,2\nfrom \"VatTuYTe\"  where \"CoBH\"=false AND \"Xoa\"=false  and \"IDVTYT\" NOT IN (select \"IDVTYT\"from \"VTYT_CanhBaoTon\" where \"MaKho\"=2)"
        ]);
    }

    public function processVTYTKiGui(Request $request)
    {
        $request->validate([
            'excel_file' => 'required|mimes:xlsx,xls,csv'
        ]);

        $file = $request->file('excel_file');
        $spreadsheet = IOFactory::load($file->getRealPath());
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        
        $queries = [];
        $maVTYTList = [];
        
        for ($row = 3; $row <= $highestRow; $row++) {
            $maNhomVTYT = $worksheet->getCell('B' . $row)->getValue();
            $maVTYT = $worksheet->getCell('E' . $row)->getValue();
            $tenVTYT = $worksheet->getCell('G' . $row)->getValue();
            $nhaSX = $worksheet->getCell('H' . $row)->getValue();
            $dvt = $worksheet->getCell('I' . $row)->getValue();
            $quiCach = $worksheet->getCell('K' . $row)->getValue();

            if (empty($maVTYT)) continue;

            $maVTYTList[] = $maVTYT;

            $tenVTYTEsc = str_replace("'", "''", $tenVTYT);
            $nhaSXEsc = str_replace("'", "''", $nhaSX);
            $dvtEsc = str_replace("'", "''", $dvt);
            $quiCachEsc = str_replace("'", "''", $quiCach);

            $query = "INSERT INTO \"VatTuYTe\"(\"MaVTYT\",\"TenVTYT\",\"NhaSX\",\"DVT\",\"QuiCach\",\"SoLuongCon\",\"CoBH\",\"MaNhomVTYT\")\n";
            $query .= "VALUES('$maVTYT',N'$tenVTYTEsc',N'$nhaSXEsc',N'$dvtEsc',N'$quiCachEsc',0,false,$maNhomVTYT);";
            
            $queries[] = $query;
        }

        $maVTYTList = array_map(function($m) { return "'$m'"; }, array_filter($maVTYTList));
        $maVTYTIn = implode(',', $maVTYTList);

        return response()->json([
            'success' => true,
            'queries' => implode("\n", $queries),
            'init_queries' => "Drop Table If Exists tbTam;\nCreate Temp Table tbTam as select \"IDVTYT\",\"MaKhoa\", cast(\"IDVTYT\" as text) || '_'|| cast(\"MaKhoa\" as text) as \"col\" from \"VatTuYTe\" as \"a\",\"Khoa\" as \"b\" ;\n\ninsert into \"SoLuongVTYTTon\"(\"IDVTYT\",\"MaKhoa\",\"SoLuong\")\nselect \"a\".\"IDVTYT\",\"a\".\"MaKhoa\",0  from tbTam as \"a\" left outer join \"SoLuongVTYTTon\" as \"b\" on \"a\".\"col\"=cast(\"b\".\"IDVTYT\" as text) || '_'|| cast(\"b\".\"MaKhoa\" as text)\nwhere \"b\".\"MaKhoa\" is null;",
            'warning_queries' => "insert into \"VTYT_CanhBaoTon\" ( \"IDVTYT\",  \"SoLuong\" ,  \"MaKho\")\nselect   \"IDVTYT\",0,1\nfrom \"VatTuYTe\"  where \"CoBH\"=false AND \"Xoa\"=false  and \"IDVTYT\" NOT IN (select \"IDVTYT\"from \"VTYT_CanhBaoTon\" where \"MaKho\"=1)",
            'update_kigui_queries' => "update \"VatTuYTe\" set \"NhaThuocNgoai\"=true,\"IDDVCTN\"=3\nwhere \"MaVTYT\" in ($maVTYTIn);",
            'check_queries' => "select  * from \"VatTuYTe\"  where \"MaVTYT\" in ($maVTYTIn);"
        ]);
    }
}

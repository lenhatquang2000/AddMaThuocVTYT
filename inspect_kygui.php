<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

$filePath = 'public/ExcelNhaThuoc/AddMaThuoc_kygui_19.05.2026_lan01.xlsx';

if (!file_exists($filePath)) {
    die("File not found: $filePath\n");
}

try {
    $spreadsheet = IOFactory::load($filePath);
    $worksheet = $spreadsheet->getActiveSheet();
    $highestColumn = $worksheet->getHighestColumn();
    
    // Đọc 5 dòng đầu tiên để xem cấu trúc
    $data = $worksheet->rangeToArray('A1:' . $highestColumn . '5', NULL, TRUE, FALSE);
    
    echo "Cấu trúc file Ký gửi:\n";
    foreach ($data as $index => $row) {
        echo "Row " . ($index + 1) . ": " . implode(' | ', array_map(function($v) { return (string)$v; }, $row)) . "\n";
    }

    echo "\nChi tiết các cột (Dòng 2):\n";
    $headers = $data[1];
    foreach ($headers as $index => $header) {
        $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($index + 1);
        $headerName = $header ? trim($header) : "[Trống]";
        echo "Cột $colLetter: $headerName\n";
    }
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
}

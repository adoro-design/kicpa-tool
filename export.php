<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
Auth::check();

$year    = (int)($_GET['year']   ?? 2026);
$dept    = $_GET['dept']         ?? '';
$month   = $_GET['month']        ?? '';
$format  = $_GET['format']       ?? '';
$billing = $_GET['billing']      ?? '';
$search  = trim($_GET['search']  ?? '');

$where  = ["year = ?"];
$params = [$year];
if ($dept)    { $where[] = "department = ?";          $params[] = $dept; }
if ($month)   { $where[] = "shooting_month = ?";      $params[] = $month; }
if ($format)  { $where[] = "shooting_format LIKE ?";  $params[] = "%$format%"; }
if ($billing === 'Y') { $where[] = "billing IS NOT NULL AND billing \!= ''"; }
if ($billing === 'N') { $where[] = "(billing IS NULL OR billing = '')"; }
if ($search)  { $where[] = "course_name LIKE ?";      $params[] = "%$search%"; }

$rows = DB::fetchAll(
    "SELECT * FROM kicpa_contents WHERE " . implode(' AND ', $where) . " ORDER BY id ASC",
    $params
);

require_once BASE_PATH . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Border;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('개발관리');

// 헤더
$sheet->mergeCells('A1:Z1');
$sheet->setCellValue('A1', "{$year}년 KICPA 콘텐츠개발 및 동영상 촬영 현황");
$sheet->getStyle('A1')->getFont()->setBold(true)->setSize(13);

$headers = [
    'No','촬영월','과정명','필수/선택','원코드','카테고리','과정코드',
    '차시수','챕터수','강사','담당부서','한공회 담당','촬영동의서',
    '촬영날짜','촬영시간','촬영형식','장소','퀴즈유무','퀴즈문항수','교안수급',
    '동영상마킹','개발(외주)','검수','오픈일','비용청구','비고'
];
foreach ($headers as $i => $h) {
    $col = chr(65 + $i);
    $sheet->setCellValue("{$col}2", $h);
    $sheet->getStyle("{$col}2")->applyFromArray([
        'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => '2F5496']],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true],
    ]);
}

// 데이터
foreach ($rows as $ri => $row) {
    $r = $ri + 3;
    $values = [
        $ri + 1,
        $row['shooting_month'],
        $row['course_name'],
        $row['required_optional'],
        $row['original_code'],
        $row['category'],
        $row['course_code'],
        $row['session_count'],
        $row['chapter_count'],
        $row['instructor'],
        $row['department'],
        $row['kicpa_manager'],
        $row['filming_consent'],
        $row['shooting_date'],
        $row['shooting_time'],
        $row['shooting_format'],
        $row['location'],
        $row['has_quiz'],
        $row['quiz_count'],
        $row['materials_supply'],
        $row['video_marking'],
        $row['dev_outsource_date'],
        $row['inspection_date'],
        $row['open_date'],
        $row['billing'],
        $row['notes'],
    ];
    foreach ($values as $ci => $val) {
        $col = chr(65 + $ci);
        $sheet->setCellValue("{$col}{$r}", $val);
        $sheet->getStyle("{$col}{$r}")->getAlignment()->setVerticalAlignment(Alignment::VERTICAL_CENTER)->setWrapText(true);
    }
    $bg = ($ri % 2 === 0) ? 'F2F7FD' : 'FFFFFF';
    $sheet->getStyle("A{$r}:Z{$r}")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB($bg);
}

// 열 너비
$widths = [6,10,45,8,10,30,25,7,7,20,16,14,10,12,14,12,18,8,8,8,10,12,10,12,12,20];
foreach ($widths as $i => $w) {
    $sheet->getColumnDimension(chr(65+$i))->setWidth($w);
}
$sheet->getRowDimension(2)->setRowHeight(30);

$filename = "{$year}_콘텐츠개발_개발현황_" . date('Ymd') . ".xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . rawurlencode($filename) . '"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

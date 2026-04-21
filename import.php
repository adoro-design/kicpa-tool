<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::check();

$msg = '';
$msg_type = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excel_file'])) {
    $file = $_FILES['excel_file'];
    $year = (int)($_POST['year'] ?? 2026);
    $mode = $_POST['import_mode'] ?? 'append'; // append | replace

    if ($file['error'] === 0 && in_array(pathinfo($file['name'], PATHINFO_EXTENSION), ['xlsx','xls'])) {
        $tmp = UPLOAD_PATH . '/import_' . time() . '.xlsx';
        move_uploaded_file($file['tmp_name'], $tmp);

        // Python으로 파싱 (서버에서 사용 불가 시 PhpSpreadsheet 사용)
        // 여기서는 PhpSpreadsheet 방식 사용
        require_once BASE_PATH . '/vendor/autoload.php';
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp);
            $sheet = $spreadsheet->getSheetByName('개발관리') ?? $spreadsheet->getActiveSheet();

            if ($mode === 'replace') {
                DB::query("DELETE FROM kicpa_contents WHERE year = ?", [$year]);
            }

            $imported = 0;
            $prev_month = '';
            foreach ($sheet->getRowIterator(4) as $row) {
                $cells = [];
                foreach ($row->getCellIterator('A', 'Z') as $cell) {
                    $val = $cell->getValue();
                    if ($val instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                        $val = $val->getPlainText();
                    }
                    $cells[] = $val;
                }

                $course_name = trim($cells[2] ?? '');
                if (empty($course_name)) continue;

                // 촬영월 병합셀 처리
                $month_val = trim($cells[1] ?? '');
                if (\!empty($month_val)) $prev_month = $month_val;
                $shooting_month = $prev_month;

                // 날짜 변환
                $to_date = function($val): ?string {
                    if (empty($val)) return null;
                    if (is_numeric($val)) {
                        return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($val)->format('Y-m-d');
                    }
                    $d = date_create($val);
                    return $d ? date_format($d, 'Y-m-d') : null;
                };

                DB::query(
                    "INSERT INTO kicpa_contents
                     (year, shooting_month, course_name, required_optional, original_code, category,
                      course_code, session_count, chapter_count, instructor, department, kicpa_manager,
                      filming_consent, shooting_date, shooting_time, shooting_format, location,
                      has_quiz, quiz_count, materials_supply, video_marking, dev_outsource_date,
                      inspection_date, open_date, billing, notes)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [
                        $year,
                        $shooting_month,
                        $course_name,
                        $cells[3] ?? null,
                        $cells[4] ?? null,
                        $cells[5] ?? null,
                        $cells[6] ?? null,
                        is_numeric($cells[7] ?? '') ? (int)$cells[7] : null,
                        is_numeric($cells[8] ?? '') ? (int)$cells[8] : null,
                        $cells[9] ?? null,
                        $cells[10] ?? null,
                        $cells[11] ?? null,
                        $cells[12] ?? null,
                        $to_date($cells[13] ?? null),
                        $cells[14] ?? null,
                        $cells[15] ?? null,
                        $cells[16] ?? null,
                        $cells[17] ?? null,
                        is_numeric($cells[18] ?? '') ? (int)$cells[18] : null,
                        $cells[19] ?? null,
                        $cells[20] ?? null,
                        $to_date($cells[21] ?? null),
                        $to_date($cells[22] ?? null),
                        $to_date($cells[23] ?? null),
                        $cells[24] ?? null,
                        $cells[25] ?? null,
                    ]
                );
                $imported++;
            }
            unlink($tmp);
            $msg = "{$imported}건의 콘텐츠를 성공적으로 가져왔습니다.";
            $msg_type = 'success';
        } catch (Exception $e) {
            $msg = '가져오기 오류: ' . $e->getMessage();
            $msg_type = 'danger';
        }
    } else {
        $msg = 'xlsx 파일만 업로드 가능합니다.';
        $msg_type = 'danger';
    }
}

render_head('Excel 가져오기');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('Excel 가져오기'); ?>
<div class="content-area">
  <?php if ($msg): ?>
  <div class="alert alert-<?= $msg_type ?>"><?= h($msg) ?></div>
  <?php endif; ?>

  <div class="row g-3">
    <div class="col-md-6">
      <div class="table-card p-4">
        <h6 class="fw-bold mb-3"><i class="bi-upload text-primary"></i> Excel 파일 업로드</h6>
        <form method="POST" enctype="multipart/form-data">
          <div class="mb-3">
            <label class="form-label">연도</label>
            <select name="year" class="form-select form-select-sm">
              <?php foreach ([2024,2025,2026,2027] as $y): ?>
              <option value="<?= $y ?>" <?= $y===2026?'selected':'' ?>><?= $y ?>년</option>
              <?php endforeach; ?>
            </select>
          </div>
          <div class="mb-3">
            <label class="form-label">가져오기 방식</label>
            <div>
              <div class="form-check form-check-inline">
                <input class="form-check-input" type="radio" name="import_mode" value="append" checked id="m1">
                <label class="form-check-label" for="m1">추가 (기존 데이터 유지)</label>
              </div>
              <div class="form-check form-check-inline">
                <input class="form-check-input" type="radio" name="import_mode" value="replace" id="m2">
                <label class="form-check-label text-danger" for="m2">덮어쓰기 (기존 데이터 삭제 후 재등록)</label>
              </div>
            </div>
          </div>
          <div class="mb-4">
            <label class="form-label">Excel 파일 선택</label>
            <input type="file" name="excel_file" class="form-control form-control-sm" accept=".xlsx,.xls" required>
            <small class="text-muted">2026_컨텐츠개발_촬영_개발현황.xlsx 형식의 파일을 업로드하세요.</small>
          </div>
          <button type="submit" class="btn btn-primary"><i class="bi-upload"></i> 가져오기 실행</button>
        </form>
      </div>
    </div>
    <div class="col-md-6">
      <div class="table-card p-4">
        <h6 class="fw-bold mb-3"><i class="bi-info-circle text-info"></i> 업로드 안내</h6>
        <ul class="text-muted" style="font-size:13px; line-height:1.9">
          <li>Excel 파일은 <code>개발관리</code> 시트 기준으로 읽습니다.</li>
          <li>데이터는 4행부터 시작합니다.</li>
          <li>병합된 촬영월 셀은 자동으로 채워집니다.</li>
          <li>날짜 형식은 자동으로 변환됩니다.</li>
          <li><strong class="text-danger">덮어쓰기</strong>를 선택하면 해당 연도의 기존 데이터가 모두 삭제됩니다.</li>
        </ul>
      </div>
    </div>
  </div>
</div>
</div>
</div>
<?php render_foot(); ?>

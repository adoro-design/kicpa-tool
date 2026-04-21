<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::check();

$id  = (int)($_GET['id'] ?? 0);
$row = $id ? DB::fetch("SELECT * FROM kicpa_contents WHERE id=?", [$id]) : null;
if ($id && \!$row) { header('Location: contents.php'); exit; }

$msg = ''; $msg_type = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $fields = [
        'shooting_month','course_name','required_optional','original_code','category',
        'course_code','session_count','chapter_count','instructor','department',
        'kicpa_manager','filming_consent','shooting_date','shooting_time','shooting_format',
        'location','has_quiz','quiz_count','materials_supply','video_marking',
        'dev_outsource_date','inspection_date','open_date','billing','notes'
    ];
    $set = implode(', ', array_map(fn($f) => "`$f`=?", $fields));
    $vals = array_map(function($f) {
        $v = $_POST[$f] ?? null;
        return ($v === '') ? null : $v;
    }, $fields);

    if ($id) {
        DB::query("UPDATE kicpa_contents SET $set WHERE id=?", array_merge($vals, [$id]));
        $msg = '저장되었습니다.'; $msg_type = 'success';
        $row = DB::fetch("SELECT * FROM kicpa_contents WHERE id=?", [$id]);
    } else {
        $year = (int)($_POST['year'] ?? 2026);
        DB::query("INSERT INTO kicpa_contents (year, $set) VALUES (?," . implode(',', array_fill(0, count($fields), '?')) . ")",
            array_merge([$year], $vals));
        $new_id = DB::lastInsertId();
        header('Location: content_edit.php?id=' . $new_id . '&saved=1'); exit;
    }
}

$v = $row ?? [];
$def = fn($k, $d='') => $v[$k] ?? $d;

render_head($id ? '콘텐츠 수정' : '콘텐츠 등록');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar($id ? '콘텐츠 수정' : '콘텐츠 등록'); ?>
<div class="content-area">
  <?php if ($msg): ?><div class="alert alert-<?= $msg_type ?>"><?= h($msg) ?></div><?php endif; ?>
  <?php if (isset($_GET['saved'])): ?><div class="alert alert-success">등록되었습니다.</div><?php endif; ?>

  <form method="POST">
    <div class="row g-3">

      <\!-- 기본 정보 -->
      <div class="col-12">
        <div class="table-card p-4">
          <h6 class="fw-bold mb-3 text-primary">기본 정보</h6>
          <div class="row g-3">
            <div class="col-md-2">
              <label class="form-label">촬영월</label>
              <select name="shooting_month" class="form-select form-select-sm">
                <option value="">선택</option>
                <?php foreach (get_months() as $m): ?>
                <option value="<?= $m ?>" <?= $def('shooting_month')===$m?'selected':'' ?>><?= $m ?></option>
                <?php endforeach; ?>
              </select>
            </div>
            <div class="col-md-2">
              <label class="form-label">필수/선택</label>
              <select name="required_optional" class="form-select form-select-sm">
                <option value="">선택</option>
                <option value="필수" <?= $def('required_optional')==='필수'?'selected':'' ?>>필수</option>
                <option value="선택" <?= $def('required_optional')==='선택'?'selected':'' ?>>선택</option>
              </select>
            </div>
            <div class="col-md-8">
              <label class="form-label">과정명</label>
              <input type="text" name="course_name" class="form-control form-control-sm" value="<?= h($def('course_name')) ?>" required>
            </div>
            <div class="col-md-3">
              <label class="form-label">담당부서</label>
              <input type="text" name="department" class="form-control form-control-sm" value="<?= h($def('department')) ?>">
            </div>
            <div class="col-md-3">
              <label class="form-label">한공회 담당</label>
              <input type="text" name="kicpa_manager" class="form-control form-control-sm" value="<?= h($def('kicpa_manager')) ?>">
            </div>
            <div class="col-md-3">
              <label class="form-label">강사</label>
              <input type="text" name="instructor" class="form-control form-control-sm" value="<?= h($def('instructor')) ?>">
            </div>
            <div class="col-md-1">
              <label class="form-label">차시수</label>
              <input type="number" name="session_count" class="form-control form-control-sm" value="<?= h($def('session_count')) ?>">
            </div>
            <div class="col-md-1">
              <label class="form-label">챕터수</label>
              <input type="number" name="chapter_count" class="form-control form-control-sm" value="<?= h($def('chapter_count')) ?>">
            </div>
            <div class="col-md-2">
              <label class="form-label">원코드</label>
              <input type="text" name="original_code" class="form-control form-control-sm" value="<?= h($def('original_code')) ?>">
            </div>
            <div class="col-md-4">
              <label class="form-label">과정코드</label>
              <input type="text" name="course_code" class="form-control form-control-sm" value="<?= h($def('course_code')) ?>">
            </div>
            <div class="col-md-8">
              <label class="form-label">카테고리</label>
              <input type="text" name="category" class="form-control form-control-sm" value="<?= h($def('category')) ?>">
            </div>
          </div>
        </div>
      </div>

      <\!-- 촬영 정보 -->
      <div class="col-md-6">
        <div class="table-card p-4 h-100">
          <h6 class="fw-bold mb-3 text-primary">촬영 정보</h6>
          <div class="row g-3">
            <div class="col-6">
              <label class="form-label">촬영날짜</label>
              <input type="date" name="shooting_date" class="form-control form-control-sm" value="<?= h($def('shooting_date')) ?>">
            </div>
            <div class="col-6">
              <label class="form-label">촬영시간</label>
              <input type="text" name="shooting_time" class="form-control form-control-sm" value="<?= h($def('shooting_time')) ?>" placeholder="10:00~12:00">
            </div>
            <div class="col-6">
              <label class="form-label">촬영형식</label>
              <select name="shooting_format" class="form-select form-select-sm">
                <option value="">선택</option>
                <?php foreach (['크로마키','FullVod (출장)','태블릿형','전자칠판형'] as $f): ?>
                <option value="<?= $f ?>" <?= $def('shooting_format')===$f?'selected':'' ?>><?= $f ?></option>
                <?php endforeach; ?>
              </select>
            </div>
            <div class="col-6">
              <label class="form-label">장소</label>
              <input type="text" name="location" class="form-control form-control-sm" value="<?= h($def('location')) ?>">
            </div>
            <div class="col-4">
              <label class="form-label">촬영 동의서</label>
              <select name="filming_consent" class="form-select form-select-sm">
                <option value="">-</option>
                <option value="●" <?= $def('filming_consent')==='●'?'selected':'' ?>>완료(●)</option>
              </select>
            </div>
            <div class="col-4">
              <label class="form-label">퀴즈 유무</label>
              <select name="has_quiz" class="form-select form-select-sm">
                <option value="">없음</option>
                <option value="●" <?= $def('has_quiz')==='●'?'selected':'' ?>>있음(●)</option>
              </select>
            </div>
            <div class="col-4">
              <label class="form-label">퀴즈 문항수</label>
              <input type="number" name="quiz_count" class="form-control form-control-sm" value="<?= h($def('quiz_count')) ?>">
            </div>
          </div>
        </div>
      </div>

      <\!-- 개발/오픈 일정 -->
      <div class="col-md-6">
        <div class="table-card p-4 h-100">
          <h6 class="fw-bold mb-3 text-primary">개발 · 오픈 일정</h6>
          <div class="row g-3">
            <div class="col-6">
              <label class="form-label">교안수급</label>
              <select name="materials_supply" class="form-select form-select-sm">
                <option value="">-</option>
                <option value="●" <?= $def('materials_supply')==='●'?'selected':'' ?>>완료(●)</option>
              </select>
            </div>
            <div class="col-6">
              <label class="form-label">동영상 마킹값</label>
              <select name="video_marking" class="form-select form-select-sm">
                <option value="">-</option>
                <option value="●" <?= $def('video_marking')==='●'?'selected':'' ?>>완료(●)</option>
              </select>
            </div>
            <div class="col-6">
              <label class="form-label">개발(외주)일</label>
              <input type="date" name="dev_outsource_date" class="form-control form-control-sm" value="<?= h($def('dev_outsource_date')) ?>">
            </div>
            <div class="col-6">
              <label class="form-label">검수일</label>
              <input type="date" name="inspection_date" class="form-control form-control-sm" value="<?= h($def('inspection_date')) ?>">
            </div>
            <div class="col-6">
              <label class="form-label">오픈일</label>
              <input type="date" name="open_date" class="form-control form-control-sm" value="<?= h($def('open_date')) ?>">
            </div>
            <div class="col-6">
              <label class="form-label">비용청구</label>
              <input type="text" name="billing" class="form-control form-control-sm" value="<?= h($def('billing')) ?>" placeholder="예: 1월 청구">
            </div>
            <div class="col-12">
              <label class="form-label">비고</label>
              <textarea name="notes" class="form-control form-control-sm" rows="3"><?= h($def('notes')) ?></textarea>
            </div>
          </div>
        </div>
      </div>

      <\!-- 버튼 -->
      <div class="col-12 d-flex gap-2">
        <button type="submit" class="btn btn-primary"><i class="bi-save"></i> 저장</button>
        <a href="contents.php" class="btn btn-outline-secondary">목록으로</a>
      </div>
    </div>
  </form>
</div>
</div>
</div>
<?php render_foot(); ?>

<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::check();

$year  = (int)($_GET['year']  ?? 2026);
$month = $_GET['month']       ?? '';
$dept  = $_GET['dept']        ?? '';

$where  = ["year = ?", "shooting_date IS NOT NULL"];
$params = [$year];
if ($month) { $where[] = "shooting_month = ?"; $params[] = $month; }
if ($dept)  { $where[] = "department = ?";     $params[] = $dept; }

$rows = DB::fetchAll(
    "SELECT * FROM kicpa_contents WHERE " . implode(' AND ', $where) . " ORDER BY shooting_date ASC",
    $params
);

$depts = DB::fetchAll("SELECT DISTINCT department FROM kicpa_contents WHERE year=? AND department IS NOT NULL ORDER BY department", [$year]);

render_head('촬영 일정');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('촬영 일정'); ?>
<div class="content-area">

  <\!-- 필터 -->
  <form method="GET" class="filter-bar mb-4">
    <input type="hidden" name="year" value="<?= $year ?>">
    <div class="d-flex gap-2 align-items-center flex-wrap">
      <select name="month" class="form-select form-select-sm" style="width:auto">
        <option value="">전체 월</option>
        <?php foreach (get_months() as $m): ?>
        <option value="<?= $m ?>" <?= $month===$m?'selected':'' ?>><?= $m ?></option>
        <?php endforeach; ?>
      </select>
      <select name="dept" class="form-select form-select-sm" style="width:auto">
        <option value="">전체 부서</option>
        <?php foreach ($depts as $d): ?>
        <option value="<?= h($d['department']) ?>" <?= $dept===$d['department']?'selected':'' ?>><?= h($d['department']) ?></option>
        <?php endforeach; ?>
      </select>
      <button type="submit" class="btn btn-primary btn-sm">조회</button>
      <a href="schedule.php?year=<?= $year ?>" class="btn btn-outline-secondary btn-sm">초기화</a>
    </div>
  </form>

  <\!-- 일정 카드 목록 -->
  <?php
  // 날짜별로 그룹핑
  $grouped = [];
  foreach ($rows as $r) {
      $grouped[$r['shooting_date']][] = $r;
  }
  ?>

  <?php if (empty($grouped)): ?>
  <div class="text-center py-5 text-muted">조회된 촬영 일정이 없습니다.</div>
  <?php endif; ?>

  <?php foreach ($grouped as $date => $items): ?>
  <div class="mb-4">
    <div class="d-flex align-items-center gap-2 mb-2">
      <span class="fw-bold" style="font-size:15px; color:#1e2a3b;"><?= date('Y년 m월 d일 (D)', strtotime($date)) ?></span>
      <span class="badge bg-primary rounded-pill"><?= count($items) ?>건</span>
    </div>
    <div class="row g-2">
    <?php foreach ($items as $r): ?>
    <div class="col-md-6 col-lg-4">
      <div class="card border-0 shadow-sm h-100" style="border-left: 4px solid #1a56db \!important;">
        <div class="card-body p-3">
          <div class="d-flex justify-content-between align-items-start mb-1">
            <?= format_badge($r['shooting_format'] ?? '') ?>
            <small class="text-muted fw-bold"><?= h($r['shooting_time'] ?? '') ?></small>
          </div>
          <div class="fw-bold mb-1" style="font-size:13px; line-height:1.4">
            <?= h(preg_replace('/\[.*?\]/', '', explode("\n", $r['course_name'] ?? '')[0])) ?>
          </div>
          <div class="text-muted" style="font-size:12px">
            <i class="bi-person"></i> <?= h($r['instructor'] ?? '-') ?>
            &nbsp;|&nbsp;
            <i class="bi-building"></i> <?= h($r['department'] ?? '-') ?>
          </div>
          <?php if ($r['location']): ?>
          <div class="text-muted mt-1" style="font-size:12px"><i class="bi-geo-alt"></i> <?= h($r['location']) ?></div>
          <?php endif; ?>
          <div class="mt-1" style="font-size:12px">
            <?= $r['session_count'] ?>차시 / <?= $r['chapter_count'] ?>챕터
            <?php if ($r['has_quiz'] === '●'): ?>
            &nbsp;<span class="badge bg-warning text-dark">퀴즈</span>
            <?php endif; ?>
          </div>
        </div>
      </div>
    </div>
    <?php endforeach; ?>
    </div>
  </div>
  <?php endforeach; ?>

</div>
</div>
</div>
<?php render_foot(); ?>

<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::check();

$year = (int)($_GET['year'] ?? 2026);

// 통계
$total     = DB::fetch("SELECT COUNT(*) as c FROM kicpa_contents WHERE year=?", [$year])['c'];
$shot      = DB::fetch("SELECT COUNT(*) as c FROM kicpa_contents WHERE year=? AND shooting_date IS NOT NULL", [$year])['c'];
$opened    = DB::fetch("SELECT COUNT(*) as c FROM kicpa_contents WHERE year=? AND open_date IS NOT NULL", [$year])['c'];
$billed    = DB::fetch("SELECT COUNT(*) as c FROM kicpa_contents WHERE year=? AND billing IS NOT NULL AND billing \!= ''", [$year])['c'];

// 부서별 현황
$dept_stats = DB::fetchAll(
    "SELECT department, COUNT(*) as total,
            SUM(CASE WHEN shooting_date IS NOT NULL THEN 1 ELSE 0 END) as shot,
            SUM(CASE WHEN open_date IS NOT NULL THEN 1 ELSE 0 END) as opened
     FROM kicpa_contents WHERE year=? AND department IS NOT NULL
     GROUP BY department ORDER BY total DESC",
    [$year]
);

// 월별 촬영 현황
$monthly = DB::fetchAll(
    "SELECT shooting_month, COUNT(*) as cnt
     FROM kicpa_contents WHERE year=? AND shooting_month IS NOT NULL
     GROUP BY shooting_month ORDER BY FIELD(shooting_month,'1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월')",
    [$year]
);

// 최근 등록 콘텐츠
$recent = DB::fetchAll(
    "SELECT course_name, department, shooting_format, shooting_date, billing
     FROM kicpa_contents WHERE year=? ORDER BY id DESC LIMIT 8",
    [$year]
);

render_head('대시보드');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar("{$year}년 콘텐츠 개발 현황"); ?>
<div class="content-area">

  <\!-- 연도 선택 -->
  <div class="d-flex gap-2 mb-4">
    <?php foreach ([2025,2026,2027] as $y): ?>
    <a href="?year=<?= $y ?>" class="btn btn-sm <?= $y===$year ? 'btn-primary' : 'btn-outline-secondary' ?>"><?= $y ?>년</a>
    <?php endforeach; ?>
  </div>

  <\!-- 통계 카드 -->
  <div class="row g-3 mb-4">
    <?php
    $stats = [
        ['icon'=>'bi-collection','label'=>'전체 콘텐츠','val'=>$total,  'color'=>'#1a56db'],
        ['icon'=>'bi-camera-video','label'=>'촬영 완료',  'val'=>$shot,  'color'=>'#0891b2'],
        ['icon'=>'bi-play-circle','label'=>'오픈 완료',   'val'=>$opened,'color'=>'#059669'],
        ['icon'=>'bi-receipt',   'label'=>'비용 청구',    'val'=>$billed,'color'=>'#d97706'],
    ];
    foreach ($stats as $s): ?>
    <div class="col-6 col-md-3">
      <div class="stat-card d-flex justify-content-between align-items-center">
        <div>
          <div class="stat-num" style="color:<?= $s['color'] ?>"><?= number_format($s['val']) ?></div>
          <div class="stat-label"><?= $s['label'] ?></div>
        </div>
        <i class="<?= $s['icon'] ?> stat-icon" style="color:<?= $s['color'] ?>"></i>
      </div>
    </div>
    <?php endforeach; ?>
  </div>

  <div class="row g-3">
    <\!-- 부서별 현황 -->
    <div class="col-md-7">
      <div class="table-card">
        <div class="p-3 border-bottom fw-bold" style="font-size:14px;">&#128196; 부서별 현황</div>
        <table class="table table-hover">
          <thead><tr><th>담당부서</th><th class="text-center">전체</th><th class="text-center">촬영완료</th><th class="text-center">오픈</th><th class="text-center">진행률</th></tr></thead>
          <tbody>
          <?php foreach ($dept_stats as $d): ?>
          <tr>
            <td><span class="badge bg-light text-dark border"><?= h($d['department']) ?></span></td>
            <td class="text-center fw-bold"><?= $d['total'] ?></td>
            <td class="text-center"><?= $d['shot'] ?></td>
            <td class="text-center"><?= $d['opened'] ?></td>
            <td style="min-width:100px">
              <?php $pct = $d['total'] > 0 ? round($d['opened']/$d['total']*100) : 0; ?>
              <div class="progress" style="height:6px">
                <div class="progress-bar bg-success" style="width:<?= $pct ?>%"></div>
              </div>
              <small class="text-muted"><?= $pct ?>%</small>
            </td>
          </tr>
          <?php endforeach; ?>
          </tbody>
        </table>
      </div>
    </div>

    <\!-- 최근 등록 -->
    <div class="col-md-5">
      <div class="table-card">
        <div class="p-3 border-bottom d-flex justify-content-between align-items-center">
          <span class="fw-bold" style="font-size:14px;">&#128337; 최근 등록 콘텐츠</span>
          <a href="contents.php?year=<?= $year ?>" class="btn btn-sm btn-outline-primary">전체 보기</a>
        </div>
        <div style="max-height:360px;overflow-y:auto">
        <table class="table table-hover">
          <tbody>
          <?php foreach ($recent as $r): ?>
          <tr>
            <td>
              <div style="font-size:12px;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"><?= h(preg_replace('/\[.*?\]/', '', explode("\n",$r['course_name'])[0])) ?></div>
              <div class="text-muted" style="font-size:11px"><?= h($r['department']) ?> · <?= h($r['shooting_format'] ?? '-') ?></div>
            </td>
            <td class="text-end"><?= billing_badge($r['billing']) ?></td>
          </tr>
          <?php endforeach; ?>
          </tbody>
        </table>
        </div>
      </div>
    </div>
  </div>

</div>
</div>
</div>
<?php render_foot(); ?>

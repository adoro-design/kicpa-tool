<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::requireAdmin();

$year = (int)($_GET['year'] ?? 2026);
$dept = $_GET['dept'] ?? '';

$depts = DB::fetchAll("SELECT DISTINCT department FROM kicpa_contents WHERE year=? AND department IS NOT NULL ORDER BY department", [$year]);

// 부서별 콘텐츠 요약 (정산 문서 생성용)
$dept_summary = DB::fetchAll(
    "SELECT department,
            COUNT(*) as total,
            SUM(session_count) as total_sessions,
            SUM(chapter_count) as total_chapters,
            GROUP_CONCAT(DISTINCT shooting_month ORDER BY shooting_month) as months
     FROM kicpa_contents
     WHERE year=? AND department IS NOT NULL
     " . ($dept ? "AND department=?" : "") . "
     GROUP BY department ORDER BY department",
    $dept ? [$year, $dept] : [$year]
);

// 생성 이력
$history = DB::fetchAll(
    "SELECT d.*, u.name as creator_name
     FROM kicpa_documents d
     LEFT JOIN kicpa_users u ON d.created_by = u.id
     ORDER BY d.created_at DESC LIMIT 20"
);

$doc_types = [
    'request'    => ['label'=>'개발요청서',   'icon'=>'bi-file-earmark-word', 'color'=>'primary'],
    'attachment' => ['label'=>'첨부자료',     'icon'=>'bi-file-earmark-excel','color'=>'success'],
    'profit'     => ['label'=>'손익분석서',   'icon'=>'bi-graph-up',          'color'=>'info'],
    'profile'    => ['label'=>'프로젝트 프로파일','icon'=>'bi-file-earmark-text','color'=>'warning'],
    'studio'     => ['label'=>'스튜디오 대관료','icon'=>'bi-camera-video',    'color'=>'danger'],
];

render_head('문서 생성');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('문서 자동 생성'); ?>
<div class="content-area">

  <\!-- 필터 -->
  <form method="GET" class="filter-bar mb-4">
    <input type="hidden" name="year" value="<?= $year ?>">
    <div class="d-flex gap-2 align-items-center">
      <select name="dept" class="form-select form-select-sm" style="width:auto">
        <option value="">전체 부서</option>
        <?php foreach ($depts as $d): ?>
        <option value="<?= h($d['department']) ?>" <?= $dept===$d['department']?'selected':'' ?>><?= h($d['department']) ?></option>
        <?php endforeach; ?>
      </select>
      <button type="submit" class="btn btn-primary btn-sm">조회</button>
    </div>
  </form>

  <div class="row g-3">
    <\!-- 부서별 문서 생성 -->
    <div class="col-md-8">
      <?php foreach ($dept_summary as $ds): ?>
      <div class="table-card p-4 mb-3">
        <div class="d-flex justify-content-between align-items-start mb-3">
          <div>
            <h6 class="fw-bold mb-0"><?= h($ds['department']) ?></h6>
            <small class="text-muted"><?= $ds['total'] ?>개 과정 · <?= $ds['total_sessions'] ?>차시 · <?= $ds['total_chapters'] ?>챕터 · 촬영월: <?= h($ds['months']) ?></small>
          </div>
        </div>
        <div class="d-flex gap-2 flex-wrap">
          <?php foreach ($doc_types as $type => $info): ?>
          <?php if ($type === 'studio'): ?>
          <form method="POST" action="api/generate_doc.php" target="_blank">
            <input type="hidden" name="type" value="<?= $type ?>">
            <input type="hidden" name="dept" value="<?= h($ds['department']) ?>">
            <input type="hidden" name="year" value="<?= $year ?>">
            <button class="btn btn-sm btn-outline-<?= $info['color'] ?>">
              <i class="<?= $info['icon'] ?>"></i> <?= $info['label'] ?>
            </button>
          </form>
          <?php else: ?>
          <form method="POST" action="api/generate_doc.php" target="_blank">
            <input type="hidden" name="type" value="<?= $type ?>">
            <input type="hidden" name="dept" value="<?= h($ds['department']) ?>">
            <input type="hidden" name="year" value="<?= $year ?>">
            <button class="btn btn-sm btn-outline-<?= $info['color'] ?>">
              <i class="<?= $info['icon'] ?>"></i> <?= $info['label'] ?>
            </button>
          </form>
          <?php endif; ?>
          <?php endforeach; ?>
          <\!-- 전체 생성 -->
          <form method="POST" action="api/generate_doc.php">
            <input type="hidden" name="type" value="all">
            <input type="hidden" name="dept" value="<?= h($ds['department']) ?>">
            <input type="hidden" name="year" value="<?= $year ?>">
            <button class="btn btn-sm btn-dark">
              <i class="bi-cloud-download"></i> 전체 생성
            </button>
          </form>
        </div>
      </div>
      <?php endforeach; ?>
    </div>

    <\!-- 최근 생성 이력 -->
    <div class="col-md-4">
      <div class="table-card">
        <div class="p-3 border-bottom fw-bold" style="font-size:14px;">&#128196; 최근 생성 이력</div>
        <?php if (empty($history)): ?>
        <div class="p-3 text-muted text-center" style="font-size:13px;">생성된 문서가 없습니다.</div>
        <?php else: ?>
        <div style="max-height:500px;overflow-y:auto">
        <table class="table table-hover">
          <tbody>
          <?php foreach ($history as $h_item): ?>
          <tr>
            <td>
              <div style="font-size:12px;font-weight:600"><?= h($doc_types[$h_item['doc_type']]['label'] ?? $h_item['doc_type']) ?></div>
              <div class="text-muted" style="font-size:11px"><?= h($h_item['department']) ?> · <?= h($h_item['period'] ?? '') ?></div>
              <div class="text-muted" style="font-size:11px"><?= date('m/d H:i', strtotime($h_item['created_at'])) ?> · <?= h($h_item['creator_name'] ?? '') ?></div>
            </td>
            <td class="text-end align-middle">
              <a href="api/download_doc.php?id=<?= $h_item['id'] ?>" class="btn btn-sm btn-outline-secondary"><i class="bi-download"></i></a>
            </td>
          </tr>
          <?php endforeach; ?>
          </tbody>
        </table>
        </div>
        <?php endif; ?>
      </div>
    </div>
  </div>

</div>
</div>
</div>
<?php render_foot(); ?>

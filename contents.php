<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::check();

$year    = (int)($_GET['year']   ?? 2026);
$dept    = $_GET['dept']         ?? '';
$month   = $_GET['month']        ?? '';
$format  = $_GET['format']       ?? '';
$billing = $_GET['billing']      ?? '';
$search  = trim($_GET['search']  ?? '');
$page    = max(1, (int)($_GET['page'] ?? 1));
$per     = 20;

// WHERE 조건 구성
$where = ["c.year = ?"];
$params = [$year];
if ($dept)    { $where[] = "c.department = ?";      $params[] = $dept; }
if ($month)   { $where[] = "c.shooting_month = ?";  $params[] = $month; }
if ($format)  { $where[] = "c.shooting_format LIKE ?"; $params[] = "%$format%"; }
if ($billing === 'Y') { $where[] = "c.billing IS NOT NULL AND c.billing \!= ''"; }
if ($billing === 'N') { $where[] = "(c.billing IS NULL OR c.billing = '')"; }
if ($search)  { $where[] = "c.course_name LIKE ?";  $params[] = "%$search%"; }

$sql_where = implode(' AND ', $where);
$total_row = DB::fetch("SELECT COUNT(*) as c FROM kicpa_contents c WHERE $sql_where", $params)['c'];
$paging    = paginate($total_row, $per, $page);

$rows = DB::fetchAll(
    "SELECT * FROM kicpa_contents c WHERE $sql_where ORDER BY c.id ASC LIMIT {$per} OFFSET {$paging['offset']}",
    $params
);

// 필터 옵션
$depts   = DB::fetchAll("SELECT DISTINCT department FROM kicpa_contents WHERE year=? AND department IS NOT NULL ORDER BY department", [$year]);
$formats = DB::fetchAll("SELECT DISTINCT shooting_format FROM kicpa_contents WHERE year=? AND shooting_format IS NOT NULL ORDER BY shooting_format", [$year]);

render_head('콘텐츠 목록');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('콘텐츠 목록'); ?>
<div class="content-area">

  <\!-- 필터 -->
  <form method="GET" class="filter-bar">
    <input type="hidden" name="year" value="<?= $year ?>">
    <div class="row g-2 align-items-end">
      <div class="col-md-3">
        <input type="text" name="search" class="form-control form-control-sm" placeholder="&#128269; 과정명 검색" value="<?= h($search) ?>">
      </div>
      <div class="col-auto">
        <select name="dept" class="form-select form-select-sm">
          <option value="">전체 부서</option>
          <?php foreach ($depts as $d): ?>
          <option value="<?= h($d['department']) ?>" <?= $dept===$d['department']?'selected':'' ?>><?= h($d['department']) ?></option>
          <?php endforeach; ?>
        </select>
      </div>
      <div class="col-auto">
        <select name="month" class="form-select form-select-sm">
          <option value="">전체 월</option>
          <?php foreach (get_months() as $m): ?>
          <option value="<?= $m ?>" <?= $month===$m?'selected':'' ?>><?= $m ?></option>
          <?php endforeach; ?>
        </select>
      </div>
      <div class="col-auto">
        <select name="format" class="form-select form-select-sm">
          <option value="">전체 형식</option>
          <?php foreach ($formats as $f): ?>
          <option value="<?= h($f['shooting_format']) ?>" <?= $format===$f['shooting_format']?'selected':'' ?>><?= h($f['shooting_format']) ?></option>
          <?php endforeach; ?>
        </select>
      </div>
      <div class="col-auto">
        <select name="billing" class="form-select form-select-sm">
          <option value="">청구 전체</option>
          <option value="Y" <?= $billing==='Y'?'selected':'' ?>>청구 완료</option>
          <option value="N" <?= $billing==='N'?'selected':'' ?>>미청구</option>
        </select>
      </div>
      <div class="col-auto">
        <button type="submit" class="btn btn-primary btn-sm">검색</button>
        <a href="contents.php?year=<?= $year ?>" class="btn btn-outline-secondary btn-sm">초기화</a>
      </div>
      <div class="col-auto ms-auto">
        <a href="export.php?year=<?= $year ?>&dept=<?= urlencode($dept) ?>&month=<?= urlencode($month) ?>&format=<?= urlencode($format) ?>&billing=<?= urlencode($billing) ?>&search=<?= urlencode($search) ?>" class="btn btn-success btn-sm">
          <i class="bi-file-earmark-excel"></i> Excel 다운로드
        </a>
        <?php if (Auth::isAdmin()): ?>
        <a href="import.php" class="btn btn-outline-primary btn-sm"><i class="bi-upload"></i> 가져오기</a>
        <?php endif; ?>
      </div>
    </div>
  </form>

  <\!-- 결과 수 -->
  <div class="d-flex justify-content-between align-items-center mb-2">
    <small class="text-muted">총 <strong><?= number_format($total_row) ?></strong>건</small>
  </div>

  <\!-- 테이블 -->
  <div class="table-card">
    <div style="overflow-x:auto">
    <table class="table table-hover">
      <thead>
        <tr>
          <th>No</th><th>촬영월</th><th>과정명</th><th>부서</th><th>강사</th>
          <th>형식</th><th>촬영일</th><th>차시</th><th>오픈일</th><th>청구</th><th>관리</th>
        </tr>
      </thead>
      <tbody>
      <?php if (empty($rows)): ?>
      <tr><td colspan="11" class="text-center py-4 text-muted">조회된 콘텐츠가 없습니다.</td></tr>
      <?php endif; ?>
      <?php foreach ($rows as $i => $r): ?>
      <tr>
        <td class="text-muted"><?= $paging['offset'] + $i + 1 ?></td>
        <td><?= h($r['shooting_month'] ?? '') ?></td>
        <td style="max-width:280px">
          <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600" title="<?= h($r['course_name']) ?>">
            <?= h(preg_replace('/\[.*?\]/', '', explode("\n", $r['course_name'] ?? '')[0])) ?>
          </div>
          <?php if ($r['required_optional']): ?>
          <small class="badge bg-light text-dark border"><?= h($r['required_optional']) ?></small>
          <?php endif; ?>
        </td>
        <td><small><?= h($r['department'] ?? '') ?></small></td>
        <td><small><?= h($r['instructor'] ?? '') ?></small></td>
        <td><?= format_badge($r['shooting_format'] ?? '') ?></td>
        <td><small><?= fmt_date($r['shooting_date']) ?></small></td>
        <td class="text-center"><?= $r['session_count'] ?? '-' ?></td>
        <td><small><?= fmt_date($r['open_date']) ?></small></td>
        <td><?= billing_badge($r['billing']) ?></td>
        <td>
          <a href="content_edit.php?id=<?= $r['id'] ?>" class="btn btn-sm btn-outline-secondary"><i class="bi-pencil"></i></a>
        </td>
      </tr>
      <?php endforeach; ?>
      </tbody>
    </table>
    </div>
  </div>

  <\!-- 페이징 -->
  <?php if ($paging['total_pages'] > 1): ?>
  <nav class="mt-3 d-flex justify-content-center">
    <ul class="pagination pagination-sm">
      <?php for ($p = max(1, $page-4); $p <= min($paging['total_pages'], $page+4); $p++): ?>
      <li class="page-item <?= $p===$page?'active':'' ?>">
        <a class="page-link" href="?year=<?= $year ?>&dept=<?= urlencode($dept) ?>&month=<?= urlencode($month) ?>&format=<?= urlencode($format) ?>&billing=<?= urlencode($billing) ?>&search=<?= urlencode($search) ?>&page=<?= $p ?>"><?= $p ?></a>
      </li>
      <?php endfor; ?>
    </ul>
  </nav>
  <?php endif; ?>

</div>
</div>
</div>
<?php render_foot(); ?>

<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::requireAdmin();

$year  = (int)($_GET['year']  ?? 2026);
$month = $_GET['month']       ?? '';
$dept  = $_GET['dept']        ?? '';

$where  = ["year = ?"];
$params = [$year];
if ($month) { $where[] = "shooting_month = ?"; $params[] = $month; }
if ($dept)  { $where[] = "department = ?";     $params[] = $dept; }

$contents = DB::fetchAll(
    "SELECT * FROM kicpa_contents WHERE " . implode(' AND ', $where) . " ORDER BY shooting_month, department",
    $params
);

// 단가 조회 (촬영형식 → 단가)
$price_map = [];
foreach (DB::fetchAll("SELECT type_name, unit_price FROM kicpa_price_table WHERE category='new_dev'") as $p) {
    $price_map[$p['type_name']] = (int)$p['unit_price'];
}
$porting_map = [];
foreach (DB::fetchAll("SELECT type_name, unit_price FROM kicpa_price_table WHERE category='porting'") as $p) {
    $porting_map[$p['type_name']] = (int)$p['unit_price'];
}

$get_price = function(string $fmt) use ($price_map): int {
    foreach ($price_map as $k => $v) {
        if (str_contains($fmt, str_replace(' (출장)','',$k)) || str_contains($fmt,$k)) return $v;
    }
    return 0;
};

$depts  = DB::fetchAll("SELECT DISTINCT department FROM kicpa_contents WHERE year=? AND department IS NOT NULL ORDER BY department", [$year]);

// 부서별 합계 계산
$summary = [];
foreach ($contents as $r) {
    $d = $r['department'] ?? '미지정';
    if (\!isset($summary[$d])) $summary[$d] = ['count'=>0,'sessions'=>0,'total'=>0,'billed'=>0,'unbilled'=>0];
    $price = $get_price($r['shooting_format'] ?? '') * ($r['session_count'] ?? 0);
    $summary[$d]['count']++;
    $summary[$d]['sessions'] += ($r['session_count'] ?? 0);
    $summary[$d]['total'] += $price;
    if (\!empty($r['billing'])) $summary[$d]['billed'] += $price;
    else $summary[$d]['unbilled'] += $price;
}

render_head('정산 관리');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('정산 관리'); ?>
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
    </div>
  </form>

  <\!-- 부서별 정산 요약 -->
  <div class="table-card mb-4">
    <div class="p-3 border-bottom fw-bold">부서별 정산 요약</div>
    <div style="overflow-x:auto">
    <table class="table table-hover">
      <thead>
        <tr>
          <th>담당부서</th><th class="text-center">과정수</th><th class="text-center">차시수</th>
          <th class="text-end">예상 총액 (VAT별도)</th><th class="text-end">청구 완료</th><th class="text-end">미청구</th>
        </tr>
      </thead>
      <tbody>
      <?php $grand_total=0; $grand_billed=0; foreach ($summary as $dept_name => $s): $grand_total+=$s['total']; $grand_billed+=$s['billed']; ?>
      <tr>
        <td><strong><?= h($dept_name) ?></strong></td>
        <td class="text-center"><?= $s['count'] ?></td>
        <td class="text-center"><?= $s['sessions'] ?></td>
        <td class="text-end"><?= number_format($s['total']) ?>원</td>
        <td class="text-end text-success fw-bold"><?= number_format($s['billed']) ?>원</td>
        <td class="text-end text-danger"><?= number_format($s['unbilled']) ?>원</td>
      </tr>
      <?php endforeach; ?>
      </tbody>
      <tfoot class="table-light">
        <tr>
          <th colspan="3">합계</th>
          <th class="text-end"><?= number_format($grand_total) ?>원</th>
          <th class="text-end text-success"><?= number_format($grand_billed) ?>원</th>
          <th class="text-end text-danger"><?= number_format($grand_total - $grand_billed) ?>원</th>
        </tr>
      </tfoot>
    </table>
    </div>
  </div>

  <\!-- 상세 목록 -->
  <div class="table-card">
    <div class="p-3 border-bottom fw-bold">상세 내역</div>
    <div style="overflow-x:auto">
    <table class="table table-hover">
      <thead>
        <tr><th>촬영월</th><th>과정명</th><th>부서</th><th>형식</th><th class="text-center">차시</th><th class="text-end">단가</th><th class="text-end">금액</th><th>청구</th></tr>
      </thead>
      <tbody>
      <?php foreach ($contents as $r):
        $unit_price = $get_price($r['shooting_format'] ?? '');
        $amount = $unit_price * ($r['session_count'] ?? 0);
      ?>
      <tr>
        <td><?= h($r['shooting_month'] ?? '') ?></td>
        <td style="max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="<?= h($r['course_name']) ?>">
          <?= h(preg_replace('/\[.*?\]/', '', explode("\n",$r['course_name']??'')[0])) ?>
        </td>
        <td><small><?= h($r['department'] ?? '') ?></small></td>
        <td><?= format_badge($r['shooting_format'] ?? '') ?></td>
        <td class="text-center"><?= $r['session_count'] ?? '-' ?></td>
        <td class="text-end"><?= $unit_price ? number_format($unit_price) : '-' ?></td>
        <td class="text-end fw-bold"><?= $amount ? number_format($amount) : '-' ?></td>
        <td><?= billing_badge($r['billing']) ?></td>
      </tr>
      <?php endforeach; ?>
      </tbody>
    </table>
    </div>
  </div>

</div>
</div>
</div>
<?php render_foot(); ?>

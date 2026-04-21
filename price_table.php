<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::requireAdmin();

$msg = ''; $msg_type = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = $_POST['action'] ?? '';
    if ($action === 'update') {
        foreach ($_POST['price'] as $id => $price) {
            DB::query("UPDATE kicpa_price_table SET unit_price=? WHERE id=?",
                [is_numeric($price) ? (int)$price : null, (int)$id]);
        }
        $msg = '단가표가 저장되었습니다.'; $msg_type = 'success';
    } elseif ($action === 'add') {
        DB::query(
            "INSERT INTO kicpa_price_table (category, type_name, unit_price, unit, note) VALUES (?,?,?,?,?)",
            [$_POST['category'], $_POST['type_name'], $_POST['unit_price'] ?: null, $_POST['unit'], $_POST['note']]
        );
        $msg = '항목이 추가되었습니다.'; $msg_type = 'success';
    } elseif ($action === 'delete') {
        DB::query("DELETE FROM kicpa_price_table WHERE id=?", [(int)$_POST['item_id']]);
        $msg = '삭제되었습니다.'; $msg_type = 'success';
    }
}

$prices = DB::fetchAll("SELECT * FROM kicpa_price_table ORDER BY category, id");
$cat_labels = ['new_dev'=>'신규 개발 단가', 'porting'=>'포팅 단가', 'travel'=>'출장비'];

render_head('단가표 관리');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('단가표 관리'); ?>
<div class="content-area">
  <?php if ($msg): ?><div class="alert alert-<?= $msg_type ?>"><?= h($msg) ?></div><?php endif; ?>

  <form method="POST">
    <input type="hidden" name="action" value="update">
    <?php foreach ($cat_labels as $cat => $cat_label): ?>
    <div class="table-card mb-4">
      <div class="p-3 border-bottom fw-bold"><?= h($cat_label) ?></div>
      <table class="table">
        <thead><tr><th>유형</th><th>단가 (VAT별도)</th><th>기준</th><th>비고</th><th></th></tr></thead>
        <tbody>
        <?php foreach ($prices as $p): if ($p['category'] \!== $cat) continue; ?>
        <tr>
          <td class="fw-semibold"><?= h($p['type_name']) ?></td>
          <td style="width:160px">
            <input type="number" name="price[<?= $p['id'] ?>]"
                   class="form-control form-control-sm" value="<?= h($p['unit_price'] ?? '') ?>"
                   placeholder="별도 협의">
          </td>
          <td><small class="text-muted"><?= h($p['unit']) ?></small></td>
          <td><small class="text-muted"><?= h($p['note']) ?></small></td>
          <td>
            <form method="POST" style="display:inline" onsubmit="return confirm('삭제하시겠습니까?')">
              <input type="hidden" name="action" value="delete">
              <input type="hidden" name="item_id" value="<?= $p['id'] ?>">
              <button class="btn btn-sm btn-outline-danger"><i class="bi-trash"></i></button>
            </form>
          </td>
        </tr>
        <?php endforeach; ?>
        </tbody>
      </table>
    </div>
    <?php endforeach; ?>
    <button type="submit" class="btn btn-primary mb-4">단가 저장</button>
  </form>

  <\!-- 항목 추가 -->
  <div class="table-card p-4" style="max-width:500px">
    <h6 class="fw-bold mb-3">항목 추가</h6>
    <form method="POST">
      <input type="hidden" name="action" value="add">
      <div class="row g-2">
        <div class="col-md-4">
          <select name="category" class="form-select form-select-sm">
            <?php foreach ($cat_labels as $k => $v): ?>
            <option value="<?= $k ?>"><?= $v ?></option>
            <?php endforeach; ?>
          </select>
        </div>
        <div class="col-md-8">
          <input type="text" name="type_name" class="form-control form-control-sm" placeholder="유형명" required>
        </div>
        <div class="col-md-4">
          <input type="number" name="unit_price" class="form-control form-control-sm" placeholder="단가">
        </div>
        <div class="col-md-4">
          <input type="text" name="unit" class="form-control form-control-sm" placeholder="기준 (차시/챕터)">
        </div>
        <div class="col-md-4">
          <input type="text" name="note" class="form-control form-control-sm" placeholder="비고">
        </div>
        <div class="col-12">
          <button class="btn btn-outline-primary btn-sm">추가</button>
        </div>
      </div>
    </form>
  </div>

</div>
</div>
</div>
<?php render_foot(); ?>

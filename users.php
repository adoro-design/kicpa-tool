<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';
require_once 'includes/layout.php';
Auth::requireAdmin();

$msg = ''; $msg_type = '';

// 사용자 추가
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'add') {
    $uname = trim($_POST['username'] ?? '');
    $name  = trim($_POST['name'] ?? '');
    $pass  = $_POST['password'] ?? '';
    $role  = $_POST['role'] ?? 'director';
    if ($uname && $name && $pass) {
        try {
            DB::query(
                "INSERT INTO kicpa_users (username, password, name, role) VALUES (?,?,?,?)",
                [$uname, password_hash($pass, PASSWORD_DEFAULT), $name, $role]
            );
            $msg = '사용자가 추가되었습니다.'; $msg_type = 'success';
        } catch (Exception $e) {
            $msg = '이미 존재하는 아이디입니다.'; $msg_type = 'danger';
        }
    }
}

// 비밀번호 변경
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'change_pw') {
    $id   = (int)($_POST['user_id'] ?? 0);
    $pass = $_POST['new_password'] ?? '';
    if ($id && $pass) {
        DB::query("UPDATE kicpa_users SET password=? WHERE id=?", [password_hash($pass, PASSWORD_DEFAULT), $id]);
        $msg = '비밀번호가 변경되었습니다.'; $msg_type = 'success';
    }
}

// 활성/비활성
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'toggle') {
    $id = (int)($_POST['user_id'] ?? 0);
    DB::query("UPDATE kicpa_users SET is_active = 1 - is_active WHERE id=? AND id \!= ?", [$id, Auth::user()['id']]);
    $msg = '변경되었습니다.'; $msg_type = 'success';
}

$users = DB::fetchAll("SELECT * FROM kicpa_users ORDER BY role ASC, id ASC");

render_head('사용자 관리');
?>
<div class="d-flex">
<?php render_sidebar(); ?>
<div id="main-content" class="flex-grow-1">
<?php render_topbar('사용자 관리'); ?>
<div class="content-area">
  <?php if ($msg): ?><div class="alert alert-<?= $msg_type ?>"><?= h($msg) ?></div><?php endif; ?>

  <div class="row g-3">
    <div class="col-md-8">
      <div class="table-card">
        <div class="p-3 border-bottom fw-bold">사용자 목록</div>
        <table class="table table-hover">
          <thead><tr><th>이름</th><th>아이디</th><th>권한</th><th>상태</th><th>등록일</th><th>관리</th></tr></thead>
          <tbody>
          <?php foreach ($users as $u): ?>
          <tr>
            <td class="fw-semibold"><?= h($u['name']) ?></td>
            <td><code><?= h($u['username']) ?></code></td>
            <td><?= $u['role']==='admin' ? '<span class="badge bg-primary">관리자</span>' : '<span class="badge bg-secondary">촬영감독</span>' ?></td>
            <td><?= $u['is_active'] ? '<span class="badge bg-success">활성</span>' : '<span class="badge bg-light text-dark border">비활성</span>' ?></td>
            <td><small class="text-muted"><?= date('Y.m.d', strtotime($u['created_at'])) ?></small></td>
            <td>
              <button class="btn btn-sm btn-outline-secondary" data-bs-toggle="modal" data-bs-target="#pwModal" data-uid="<?= $u['id'] ?>" data-uname="<?= h($u['name']) ?>">
                PW변경
              </button>
              <?php if ($u['id'] \!= Auth::user()['id']): ?>
              <form method="POST" style="display:inline">
                <input type="hidden" name="action" value="toggle">
                <input type="hidden" name="user_id" value="<?= $u['id'] ?>">
                <button class="btn btn-sm <?= $u['is_active'] ? 'btn-outline-danger' : 'btn-outline-success' ?>"><?= $u['is_active'] ? '비활성화' : '활성화' ?></button>
              </form>
              <?php endif; ?>
            </td>
          </tr>
          <?php endforeach; ?>
          </tbody>
        </table>
      </div>
    </div>

    <div class="col-md-4">
      <div class="table-card p-4">
        <h6 class="fw-bold mb-3">사용자 추가</h6>
        <form method="POST">
          <input type="hidden" name="action" value="add">
          <div class="mb-2">
            <label class="form-label">이름</label>
            <input type="text" name="name" class="form-control form-control-sm" placeholder="실명" required>
          </div>
          <div class="mb-2">
            <label class="form-label">아이디</label>
            <input type="text" name="username" class="form-control form-control-sm" placeholder="영문/숫자" required>
          </div>
          <div class="mb-2">
            <label class="form-label">비밀번호</label>
            <input type="password" name="password" class="form-control form-control-sm" required>
          </div>
          <div class="mb-3">
            <label class="form-label">권한</label>
            <select name="role" class="form-select form-select-sm">
              <option value="director">촬영감독</option>
              <option value="admin">관리자</option>
            </select>
          </div>
          <button class="btn btn-primary btn-sm w-100">추가</button>
        </form>
      </div>
    </div>
  </div>

</div>
</div>
</div>

<\!-- 비밀번호 변경 모달 -->
<div class="modal fade" id="pwModal" tabindex="-1">
  <div class="modal-dialog modal-sm">
    <form method="POST" class="modal-content">
      <input type="hidden" name="action" value="change_pw">
      <input type="hidden" name="user_id" id="modal_uid">
      <div class="modal-header"><h6 class="modal-title">비밀번호 변경 - <span id="modal_uname"></span></h6><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <input type="password" name="new_password" class="form-control" placeholder="새 비밀번호" required>
      </div>
      <div class="modal-footer"><button type="submit" class="btn btn-primary btn-sm">변경</button></div>
    </form>
  </div>
</div>
<script>
document.getElementById('pwModal').addEventListener('show.bs.modal', function(e) {
  document.getElementById('modal_uid').value   = e.relatedTarget.dataset.uid;
  document.getElementById('modal_uname').textContent = e.relatedTarget.dataset.uname;
});
</script>
<?php render_foot(); ?>

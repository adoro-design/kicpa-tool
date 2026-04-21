<?php
require_once 'config.php';
require_once 'includes/db.php';
require_once 'includes/auth.php';
require_once 'includes/functions.php';

if (Auth::isLoggedIn()) {
    header('Location: ' . BASE_URL . '/dashboard.php'); exit;
}

$error = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $username = trim($_POST['username'] ?? '');
    $password = $_POST['password'] ?? '';
    if (Auth::login($username, $password)) {
        header('Location: ' . BASE_URL . '/dashboard.php'); exit;
    }
    $error = '아이디 또는 비밀번호가 올바르지 않습니다.';
}
?>
<\!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>로그인 - <?= APP_NAME ?></title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
<link rel="stylesheet" href="<?= BASE_URL ?>/assets/css/style.css">
</head>
<body>
<div class="login-wrap">
  <div class="login-card">
    <div class="text-center mb-4">
      <div style="font-size:40px; color:#1a56db;">&#128218;</div>
      <h5 class="fw-bold mt-2 mb-0">KICPA 콘텐츠 관리</h5>
      <small class="text-muted">한국공인회계사회</small>
    </div>
    <?php if ($error): ?>
    <div class="alert alert-danger py-2 text-center" style="font-size:13px;"><?= h($error) ?></div>
    <?php endif; ?>
    <form method="POST">
      <div class="mb-3">
        <label class="form-label fw-semibold">아이디</label>
        <input type="text" name="username" class="form-control" placeholder="아이디 입력" required autofocus>
      </div>
      <div class="mb-4">
        <label class="form-label fw-semibold">비밀번호</label>
        <input type="password" name="password" class="form-control" placeholder="비밀번호 입력" required>
      </div>
      <button type="submit" class="btn btn-primary w-100 fw-bold">로그인</button>
    </form>
    <p class="text-center text-muted mt-3 mb-0" style="font-size:11px;">© 2026 KICPA 콘텐츠 관리 시스템</p>
  </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<?php
// 공통 레이아웃 헬퍼
function render_head(string $title = ''): void {
    $full_title = ($title ? $title . ' - ' : '') . APP_NAME;
    echo <<<HTML
<\!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{$full_title}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
<link rel="stylesheet" href="<?= BASE_URL ?>/assets/css/style.css">
</head>
<body>
HTML;
}

function render_sidebar(): void {
    $role = $_SESSION['user_role'] ?? '';
    $name = $_SESSION['user_name'] ?? '';
    $is_admin = ($role === 'admin');
    $current = basename($_SERVER['PHP_SELF']);

    $menu_items = [
        ['icon' => 'bi-speedometer2', 'label' => '대시보드',      'file' => 'dashboard.php',       'role' => 'all'],
        ['icon' => 'bi-table',         'label' => '콘텐츠 목록',   'file' => 'contents.php',         'role' => 'all'],
        ['icon' => 'bi-camera-video',  'label' => '촬영 일정',     'file' => 'schedule.php',         'role' => 'all'],
        ['icon' => 'bi-upload',        'label' => 'Excel 가져오기','file' => 'import.php',           'role' => 'all'],
        ['icon' => 'bi-download',      'label' => 'Excel 내보내기','file' => 'export.php',           'role' => 'all'],
        ['icon' => 'bi-file-earmark-word', 'label' => '문서 생성', 'file' => 'documents.php',       'role' => 'admin'],
        ['icon' => 'bi-cash-stack',    'label' => '정산 관리',     'file' => 'billing.php',          'role' => 'admin'],
        ['icon' => 'bi-tag',           'label' => '단가표 관리',   'file' => 'price_table.php',      'role' => 'admin'],
        ['icon' => 'bi-people',        'label' => '사용자 관리',   'file' => 'users.php',            'role' => 'admin'],
    ];

    echo '<nav id="sidebar"><div class="sidebar-brand">';
    echo '<h6>KICPA</h6><strong>콘텐츠 관리</strong></div>';
    echo '<ul class="nav flex-column mt-2">';

    $prev_section = '';
    foreach ($menu_items as $item) {
        if ($item['role'] === 'admin' && \!$is_admin) continue;
        $section = in_array($item['file'], ['documents.php','billing.php','price_table.php','users.php'])
            ? '관리' : '메뉴';
        if ($section \!== $prev_section) {
            echo "<li class=\"nav-section\">{$section}</li>";
            $prev_section = $section;
        }
        $active = ($current === $item['file']) ? ' active' : '';
        echo "<li class=\"nav-item\">";
        echo "<a class=\"nav-link{$active}\" href=\"" . BASE_URL . "/{$item['file']}\">";
        echo "<i class=\"{$item['icon']}\"></i> {$item['label']}</a></li>";
    }

    echo '</ul>';
    echo "<div class=\"sidebar-footer\"><i class=\"bi-person-circle\"></i> {$name} ";
    echo ($is_admin ? '<span class="badge bg-primary">관리자</span>' : '<span class="badge bg-secondary">촬영감독</span>');
    echo "<br><a href=\"" . BASE_URL . "/logout.php\" class=\"text-danger text-decoration-none mt-1 d-inline-block\"><i class=\"bi-box-arrow-right\"></i> 로그아웃</a></div>";
    echo '</nav>';
}

function render_topbar(string $title): void {
    echo '<div class="top-bar"><h5 class="page-title">' . h($title) . '</h5></div>';
}

function render_foot(): void {
    echo <<<HTML
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body></html>
HTML;
}

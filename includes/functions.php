<?php
// ── 날짜 포맷 ──────────────────────────────────
function fmt_date(?string $date, string $fmt = 'Y.m.d'): string {
    if (empty($date) || $date === '0000-00-00') return '';
    return date($fmt, strtotime($date));
}

// ── 촬영 형식 → 단가 조회 ─────────────────────
function get_unit_price(string $format): int {
    $prices = DB::fetchAll(
        "SELECT type_name, unit_price FROM kicpa_price_table WHERE category = 'new_dev' AND is_active = 1"
    );
    foreach ($prices as $p) {
        if (str_contains($format, $p['type_name'])) return (int)$p['unit_price'];
    }
    return 0;
}

// ── 담당부서 목록 ─────────────────────────────
function get_departments(): array {
    return DB::fetchAll(
        "SELECT DISTINCT department FROM kicpa_contents WHERE department IS NOT NULL ORDER BY department"
    );
}

// ── 촬영월 목록 ──────────────────────────────
function get_months(): array {
    return ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
}

// ── HTML 이스케이프 ───────────────────────────
function h(?string $str): string {
    return htmlspecialchars($str ?? '', ENT_QUOTES, 'UTF-8');
}

// ── 페이징 계산 ───────────────────────────────
function paginate(int $total, int $per_page, int $current): array {
    $total_pages = (int)ceil($total / $per_page);
    return [
        'total'       => $total,
        'per_page'    => $per_page,
        'current'     => $current,
        'total_pages' => $total_pages,
        'offset'      => ($current - 1) * $per_page,
    ];
}

// ── 콘텐츠 유형 배지 색상 ─────────────────────
function format_badge(string $format): string {
    $colors = [
        '크로마키'   => 'primary',
        'FullVod'    => 'success',
        '태블릿'     => 'info',
        '전자칠판'   => 'warning',
    ];
    foreach ($colors as $key => $color) {
        if (str_contains($format, $key)) {
            return "<span class=\"badge bg-{$color}\">" . h($format) . "</span>";
        }
    }
    return "<span class=\"badge bg-secondary\">" . h($format) . "</span>";
}

// ── 비용청구 여부 배지 ────────────────────────
function billing_badge(?string $billing): string {
    if (empty($billing)) return '<span class="badge bg-light text-dark">미청구</span>';
    return '<span class="badge bg-success">' . h($billing) . '</span>';
}

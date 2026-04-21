<?php
class Auth {
    // 로그인 처리
    public static function login(string $username, string $password): bool {
        $user = DB::fetch(
            "SELECT * FROM kicpa_users WHERE username = ? AND is_active = 1",
            [$username]
        );
        if ($user && password_verify($password, $user['password'])) {
            $_SESSION['user_id']   = $user['id'];
            $_SESSION['user_name'] = $user['name'];
            $_SESSION['user_role'] = $user['role'];
            $_SESSION['username']  = $user['username'];
            return true;
        }
        return false;
    }

    // 로그아웃
    public static function logout(): void {
        session_destroy();
        header('Location: ' . BASE_URL . '/login.php');
        exit;
    }

    // 로그인 여부 확인
    public static function check(): void {
        if (empty($_SESSION['user_id'])) {
            header('Location: ' . BASE_URL . '/login.php');
            exit;
        }
    }

    // 관리자 권한 확인
    public static function requireAdmin(): void {
        self::check();
        if ($_SESSION['user_role'] \!== 'admin') {
            http_response_code(403);
            die('<p style="color:red;text-align:center;margin-top:100px;">접근 권한이 없습니다.</p>');
        }
    }

    public static function isAdmin(): bool {
        return ($_SESSION['user_role'] ?? '') === 'admin';
    }

    public static function isLoggedIn(): bool {
        return \!empty($_SESSION['user_id']);
    }

    public static function user(): array {
        return [
            'id'   => $_SESSION['user_id']   ?? null,
            'name' => $_SESSION['user_name'] ?? '',
            'role' => $_SESSION['user_role'] ?? '',
        ];
    }
}

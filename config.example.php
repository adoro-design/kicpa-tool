<?php
// ============================================================
//  KICPA 콘텐츠 관리 툴 - 설정 파일
//  FTP 업로드 후 아래 DB 정보를 서버 환경에 맞게 수정하세요.
// ============================================================

define('DB_HOST', 'localhost');
define('DB_NAME', 'your_database_name');   // ← 수정 필요
define('DB_USER', 'your_db_user');         // ← 수정 필요
define('DB_PASS', 'your_db_password');     // ← 수정 필요
define('DB_CHARSET', 'utf8mb4');

define('APP_NAME', 'KICPA 콘텐츠 관리');
define('APP_VERSION', '1.0.0');
define('BASE_PATH', __DIR__);
define('BASE_URL', '/kicpa-tool');         // ← 서버 경로에 맞게 수정

// 파일 저장 경로
define('UPLOAD_PATH',    BASE_PATH . '/uploads');
define('GENERATED_PATH', BASE_PATH . '/generated');
define('TEMPLATE_PATH',  BASE_PATH . '/templates');

// 세션 설정
ini_set('session.cookie_httponly', 1);
session_start();

// 타임존
date_default_timezone_set('Asia/Seoul');

<?php
require_once 'config.php';
require_once 'includes/auth.php';
if (Auth::isLoggedIn()) {
    header('Location: ' . BASE_URL . '/dashboard.php');
} else {
    header('Location: ' . BASE_URL . '/login.php');
}
exit;

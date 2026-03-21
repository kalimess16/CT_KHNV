<?php
$_SERVER['REQUEST_METHOD'] = 'POST';
$_POST['action'] = 'export';
$_POST['payload'] = '';
require __DIR__ . '/CODE/index.php';

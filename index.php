<?php
declare(strict_types=1);

function home_h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

$menuItems = [
    [
        'label' => 'Kế hoạch',
        'children' => [
            [
                'label' => 'Nhập kế hoạch nguồn vốn tín dụng',
                'description' => 'Chức năng riêng, hiện chưa triển khai.',
                'available' => false,
                'message' => 'Đang phát triển chờ',
            ],
        ],
    ],
    [
        'label' => 'Chỉ tiêu',
        'children' => [
            [
                'label' => 'Xuất tờ trình CTTD',
                'description' => 'Mở module CHITIEU để xuất biểu mẫu DOCX.',
                'available' => true,
                'href' => 'CHITIEU/index.php?view=export',
            ],
        ],
    ],
];

$cards = [
    [
        'class' => 'card-plan',
        'kicker' => 'Kế hoạch',
        'title' => 'Nhập kế hoạch nguồn vốn tín dụng',
        'description' => 'Chức năng này là module riêng và hiện chưa triển khai.',
        'cta' => 'Đang phát triển chờ',
        'available' => false,
        'message' => 'Đang phát triển chờ',
    ],
    [
        'class' => 'card-target',
        'kicker' => 'Chỉ tiêu',
        'title' => 'Xuất tờ trình CTTD',
        'description' => 'Mở module CHITIEU để rà soát dữ liệu và xuất biểu mẫu DOCX.',
        'cta' => 'Mở chức năng',
        'available' => true,
        'href' => 'CHITIEU/index.php?view=export',
    ],
];
?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PHỤC VỤ KẾ HOẠCH NGHIỆP VỤ</title>
    <link rel="stylesheet" href="assets/home.css">
</head>
<body>
<div class="page-shell">
    <header class="page-head">
        <div class="head-copy">
            <span class="eyebrow">KHNV Workspace</span>
            <h1>PHỤC VỤ KẾ HOẠCH NGHIỆP VỤ</h1>
            <p>Trang chủ đang dùng kiểu menu ngang với dropdown và chỉ giữ lại những chức năng hiện có.</p>
        </div>

        <nav class="top-nav" aria-label="Điều hướng chính">
            <ul class="top-menu">
                <?php foreach ($menuItems as $menuItem): ?>
                    <li class="top-item">
                        <button type="button" class="top-trigger" aria-expanded="false">
                            <?= home_h($menuItem['label']) ?>
                            <span class="top-caret" aria-hidden="true"></span>
                        </button>
                        <div class="dropdown">
                            <?php foreach ($menuItem['children'] as $child): ?>
                                <?php if (!empty($child['available'])): ?>
                                    <a class="dropdown-link" href="<?= home_h((string) $child['href']) ?>">
                                <?php else: ?>
                                    <button
                                        type="button"
                                        class="dropdown-link dropdown-link-disabled js-coming-soon"
                                        data-message="<?= home_h((string) ($child['message'] ?? 'Đang phát triển chờ')) ?>"
                                    >
                                <?php endif; ?>
                                        <strong><?= home_h($child['label']) ?></strong>
                                        <small><?= home_h($child['description']) ?></small>
                                <?php if (!empty($child['available'])): ?>
                                    </a>
                                <?php else: ?>
                                    </button>
                                <?php endif; ?>
                            <?php endforeach; ?>
                        </div>
                    </li>
                <?php endforeach; ?>
            </ul>
        </nav>
    </header>

    <main class="card-grid">
        <?php foreach ($cards as $card): ?>
            <?php if (!empty($card['available'])): ?>
                <a class="feature-card <?= home_h($card['class']) ?>" href="<?= home_h((string) $card['href']) ?>">
            <?php else: ?>
                <button
                    type="button"
                    class="feature-card feature-card-disabled js-coming-soon <?= home_h($card['class']) ?>"
                    data-message="<?= home_h((string) ($card['message'] ?? 'Đang phát triển chờ')) ?>"
                >
            <?php endif; ?>
                <span class="card-kicker"><?= home_h($card['kicker']) ?></span>
                <h2><?= home_h($card['title']) ?></h2>
                <p><?= home_h($card['description']) ?></p>
                <span class="card-link"><?= home_h($card['cta']) ?></span>
            <?php if (!empty($card['available'])): ?>
                </a>
            <?php else: ?>
                </button>
            <?php endif; ?>
        <?php endforeach; ?>
    </main>
</div>

<script>
(() => {
    const items = document.querySelectorAll('.top-item');

    function closeAll(exceptItem = null) {
        items.forEach((item) => {
            if (item === exceptItem) {
                return;
            }
            item.classList.remove('is-open');
            const trigger = item.querySelector('.top-trigger');
            if (trigger instanceof HTMLButtonElement) {
                trigger.setAttribute('aria-expanded', 'false');
            }
        });
    }

    items.forEach((item) => {
        const trigger = item.querySelector('.top-trigger');
        if (!(trigger instanceof HTMLButtonElement)) {
            return;
        }

        trigger.addEventListener('click', () => {
            const isOpen = item.classList.contains('is-open');
            closeAll(item);
            item.classList.toggle('is-open', !isOpen);
            trigger.setAttribute('aria-expanded', isOpen ? 'false' : 'true');
        });
    });

    document.addEventListener('click', (event) => {
        if (!(event.target instanceof Node)) {
            return;
        }
        if (!event.target.closest('.top-nav')) {
            closeAll();
        }
    });

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            closeAll();
        }
    });

    document.querySelectorAll('.js-coming-soon').forEach((button) => {
        button.addEventListener('click', () => {
            const message = button.getAttribute('data-message') || 'Đang phát triển chờ';
            window.alert(message);
        });
    });
})();
</script>
</body>
</html>

<?php
declare(strict_types=1);

require __DIR__ . '/access_control.php';

khnv_access_redirect_localhost_to_ip();
khnv_access_enforce_client_ip();

function home_h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function home_inline_src(string $href): string
{
    return $href . (str_contains($href, '?') ? '&' : '?') . 'embedded=1';
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
                'description' => 'Hổ trợ xuất tờ trình điều chỉnh kết hoạch.',
                'available' => true,
                'href' => 'CHITIEU/index.php',
            ],
        ],
    ],
];

?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PHỤC VỤ KẾ HOẠCH NGHIỆP VỤ</title>
    <link rel="icon" type="image/png" href="iconweb.png">
    <link rel="stylesheet" href="assets/home.css">
</head>
<body>
<div class="page-shell">
    <header class="page-head">
        <button type="button" class="brand-bar brand-home-trigger" id="homeBrandTrigger" aria-label="Về trang chủ KHNV">
            <div class="brand-logo">
                <img class="brand-logo-image" src="logo.png" alt="Logo VBSP">
            </div>

            <div class="head-copy">
                <h1>PHỤC VỤ KẾ HOẠCH NGHIỆP VỤ</h1>
                <p>Trang này hổ trợ công tác cá nhân trong phòng KHNV.</p>
            </div>
        </button>

        <nav class="top-nav" aria-label="Điều hướng chính">
            <ul class="top-menu">
                <?php foreach ($menuItems as $menuItem): ?>
                    <li class="top-item">
                        <button type="button" class="top-trigger" aria-expanded="false">
                            <span><?= home_h($menuItem['label']) ?></span>
                            <span class="top-caret" aria-hidden="true"></span>
                        </button>

                        <div class="dropdown">
                            <?php foreach ($menuItem['children'] as $child): ?>
                                <?php if (!empty($child['available'])): ?>
                                    <a
                                        class="dropdown-link js-inline-feature"
                                        href="<?= home_h((string) $child['href']) ?>"
                                        data-inline-src="<?= home_h(home_inline_src((string) $child['href'])) ?>"
                                    >
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

    <main class="content-shell">
        <section class="logo-showcase" id="homePlaceholder" aria-label="Logo trang chủ">
            <img class="logo-showcase-image" src="logo.png" alt="Logo VBSP">
            <p class="logo-showcase-slogan">Thấu hiểu lòng dân - Tận tâm phục vụ</p>
        </section>

        <section class="feature-viewer" id="featureViewer" hidden aria-label="Vùng hiển thị chức năng">
            <iframe
                id="featureFrame"
                class="feature-frame"
                title="Khung nội dung chức năng"
                loading="lazy"
                src="about:blank"
            ></iframe>
        </section>
    </main>

    <footer class="page-footer" aria-label="Thông tin cuối trang">
        <span>Creat by <strong>@TinHoc_DN</strong></span>
    </footer>
</div>

<script>
(() => {
    const items = document.querySelectorAll('.top-item');
    const inlineLinks = document.querySelectorAll('.js-inline-feature');
    const homeBrandTrigger = document.getElementById('homeBrandTrigger');
    const placeholder = document.getElementById('homePlaceholder');
    const viewer = document.getElementById('featureViewer');
    const featureFrame = document.getElementById('featureFrame');
    const homePath = new URL(window.location.href).pathname;

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

    function openInlineFeature(src) {
        if (!(viewer instanceof HTMLElement) || !(featureFrame instanceof HTMLIFrameElement)) {
            return;
        }

        const currentSrc = viewer.dataset.activeSrc || '';
        if (!viewer.hidden && currentSrc === src) {
            viewer.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return;
        }

        if (placeholder instanceof HTMLElement) {
            placeholder.hidden = true;
        }

        viewer.hidden = false;
        viewer.dataset.activeSrc = src;
        featureFrame.src = src;
        syncFeatureFrameHeight();

        viewer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function resetToHomePlaceholder(shouldScroll = true) {
        if (viewer instanceof HTMLElement) {
            viewer.hidden = true;
            delete viewer.dataset.activeSrc;
        }

        if (featureFrame instanceof HTMLIFrameElement) {
            featureFrame.src = 'about:blank';
            featureFrame.style.height = '';
        }

        if (placeholder instanceof HTMLElement) {
            placeholder.hidden = false;
            if (shouldScroll) {
                placeholder.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }

    function syncFeatureFrameHeight() {
        if (!(viewer instanceof HTMLElement) || !(featureFrame instanceof HTMLIFrameElement) || viewer.hidden) {
            return;
        }

        const top = viewer.getBoundingClientRect().top;
        const available = window.innerHeight - top - 4;
        featureFrame.style.height = `${Math.max(Math.floor(available), 660)}px`;
    }

    inlineLinks.forEach((link) => {
        link.addEventListener('click', (event) => {
            event.preventDefault();

            const src = link.getAttribute('data-inline-src') || link.getAttribute('href');
            if (!src) {
                return;
            }

            closeAll();
            openInlineFeature(src);
        });
    });

    if (homeBrandTrigger instanceof HTMLButtonElement) {
        homeBrandTrigger.addEventListener('click', (event) => {
            event.preventDefault();
            closeAll();
            resetToHomePlaceholder();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }

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

    window.addEventListener('resize', () => {
        syncFeatureFrameHeight();
    });

    if (featureFrame instanceof HTMLIFrameElement) {
        featureFrame.addEventListener('load', () => {
            try {
                const frameWindow = featureFrame.contentWindow;
                const frameDocument = frameWindow?.document;
                if (!frameWindow || !frameDocument) {
                    return;
                }

                const framePath = frameWindow.location.pathname;
                if (framePath === homePath || frameDocument.querySelector('.page-head')) {
                    resetToHomePlaceholder(false);
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                    return;
                }
            } catch (error) {
                console.warn('Không thể đồng bộ nội dung iframe:', error);
            }

            syncFeatureFrameHeight();
        });
    }

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

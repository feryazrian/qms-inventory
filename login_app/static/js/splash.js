(function () {
    const MIN_SPLASH_MS = 420;
    const startedAt = Date.now();

    function setSplashVisible(visible) {
        if (!document.body) {
            return;
        }

        document.body.classList.toggle("is-splash-visible", visible);
    }

    function hideSplash(immediate) {
        const elapsed = Date.now() - startedAt;
        const waitTime = immediate ? 0 : Math.max(0, MIN_SPLASH_MS - elapsed);

        window.setTimeout(function () {
            setSplashVisible(false);
        }, waitTime);
    }

    function isInternalLink(anchor) {
        const href = anchor.getAttribute("href");

        if (!href || href.startsWith("#") || anchor.hasAttribute("download")) {
            return false;
        }

        if (anchor.target && anchor.target !== "_self") {
            return false;
        }

        try {
            const targetUrl = new URL(anchor.href, window.location.href);
            return targetUrl.origin === window.location.origin;
        } catch (error) {
            return false;
        }
    }

    document.addEventListener("DOMContentLoaded", function () {
        hideSplash(false);

        document.addEventListener("click", function (event) {
            if (
                event.defaultPrevented ||
                event.button !== 0 ||
                event.metaKey ||
                event.ctrlKey ||
                event.shiftKey ||
                event.altKey
            ) {
                return;
            }

            const anchor = event.target.closest("a[href]");
            if (!anchor || !isInternalLink(anchor)) {
                return;
            }

            setSplashVisible(true);
        });

        document.addEventListener("submit", function (event) {
            const form = event.target;
            if (!form || form.hasAttribute("data-skip-splash")) {
                return;
            }

            setSplashVisible(true);
        });
    });

    window.addEventListener("pageshow", function (event) {
        if (event.persisted) {
            hideSplash(true);
        }
    });

    window.addEventListener("beforeunload", function () {
        setSplashVisible(true);
    });
})();

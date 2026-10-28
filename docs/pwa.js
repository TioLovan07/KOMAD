'use strict';
(() => {
    const status = document.getElementById('offline-status'),
        detail = document.getElementById('offline-detail'),
        install = document.getElementById('install-app'),
        update = document.getElementById('update-app'),
        check = document.getElementById('check-update'),
        dialog = document.getElementById('offline');
    let registration = null,
        promptEvent = null,
        ready = false,
        reloadRequested = false,
        cacheVersion = '',
        hadController = !!navigator.serviceWorker?.controller,
        installedInWindow = false,
        lastPreparationFailure = '';
    const standalone = () => matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;

    function installed() {
        const yes = standalone() || installedInWindow;
        install.textContent = yes ? 'Aplikasi sudah terpasang' : 'Pasang aplikasi';
        install.disabled = yes;
        if (yes) document.getElementById('install-result').textContent = 'Aplikasi terpasang. Kesiapan offline diperiksa terpisah pada status di bawah.';
    }

    function diagnostic(code, errorName = '') {
        const box = document.getElementById('offline-diagnostics');
        box.hidden = false;
        document.getElementById('offline-diagnostic-text').value = ['KOMAD 1.3.2', 'Kode: ' + code, 'Browser: ' + ((navigator.userAgent.match(/Edg\/[\d.]+|Firefox\/[\d.]+|Chrome\/[\d.]+|Version\/[\d.]+/) || ['tidak dikenali'])[0]), 'HTTPS: ' + (location.protocol === 'https:' ? 'ya' : 'tidak'), 'Konteks aman: ' + (window.isSecureContext ? 'ya' : 'tidak'), 'API offline: ' + ('serviceWorker' in navigator ? 'tersedia' : 'tidak tersedia'), 'Di dalam bingkai: ' + (window.self !== window.top ? 'ya' : 'tidak'), 'Jendela aplikasi: ' + (standalone() ? 'ya' : 'tidak'), errorName ? 'Keterangan: ' + errorName : '', lastPreparationFailure ? 'Komponen gagal: ' + lastPreparationFailure : ''].filter(Boolean).join('\n');
    }

    function state(text, description, kind) {
        status.textContent = text;
        detail.textContent = description;
        status.dataset.state = kind;
        if (kind === 'ready') document.getElementById('offline-diagnostics').hidden = true;
    }
    document.getElementById('copy-offline-diagnostic').onclick = async () => {
        const field = document.getElementById('offline-diagnostic-text'),
            result = document.getElementById('copy-offline-result');
        try {
            await navigator.clipboard.writeText(field.value);
            result.textContent = 'Info tersalin. Tempelkan saat melaporkan kendala.';
        } catch {
            field.focus();
            field.select();
            result.textContent = 'Pilih Salin atau tekan Ctrl+C untuk menyalin info.';
        }
    };
    window.addEventListener('focus', installed);
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) installed();
    });
    matchMedia('(display-mode: standalone)').addEventListener?.('change', installed);

    function networkState() {
        if (ready) state(navigator.onLine ? 'Siap digunakan offline' : 'Sedang digunakan offline', 'Buka ikon KOMAD untuk bekerja. Internet hanya diperlukan untuk pemasangan awal dan pembaruan.', 'ready');
    }

    function showHelp() {
        const ios = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1),
            safari = /Safari/.test(navigator.userAgent) && !/Chrome|Chromium|Edg|Android/.test(navigator.userAgent);
        document.getElementById('install-browser-steps').textContent = ios ? 'Di Safari: ketuk Bagikan, pilih Tambah ke Layar Utama, lalu Tambah.' : safari ? 'Di Safari pada Mac: pilih File, lalu Tambahkan ke Dock.' : 'Di Chrome atau Edge: buka menu ⋮ / …, pilih Pasang aplikasi atau Instal halaman sebagai aplikasi. Nama menu dapat berbeda sesuai browser.';
        dialog.showModal();
    }
    installed();
    window.addEventListener('beforeinstallprompt', event => {
        event.preventDefault();
        promptEvent = event;
        installed();
    });
    window.addEventListener('appinstalled', () => {
        promptEvent = null;
        installedInWindow = true;
        installed();
        document.getElementById('install-result').textContent = 'Pemasangan selesai. Selanjutnya buka ikon KOMAD dari daftar aplikasi atau layar utama perangkat.';
    });
    install.onclick = async () => {
        if (promptEvent) {
            const event = promptEvent;
            promptEvent = null;
            await event.prompt();
            const choice = await event.userChoice;
            if (choice.outcome === 'accepted') {
                document.getElementById('install-result').textContent = 'Pemasangan diterima. Buka aplikasi dari ikon KOMAD. Kesiapan offline diperiksa terpisah.';
            } else document.getElementById('install-result').textContent = 'Pemasangan belum dilakukan. Anda tetap bisa bekerja melalui halaman ini.';
        } else showHelp();
    };
    document.getElementById('close-install-help').onclick = () => dialog.close();
    document.getElementById('install-info').onclick = showHelp;
    dialog.addEventListener('click', e => {
        if (e.target === dialog) dialog.close();
    });
    window.addEventListener('online', networkState);
    window.addEventListener('offline', networkState);
    navigator.serviceWorker?.addEventListener('message', event => {
        if (event.data?.type === 'OFFLINE_PREPARE_FAILED' && !ready) {
            lastPreparationFailure = event.data.asset + ' (' + event.data.error + ')';
            state('Komponen offline belum berhasil disimpan', 'Sambungkan internet dan gunakan Periksa kembali. Info komponen yang gagal tersedia pada rincian kendala.', 'error');
            diagnostic('ASSET_FAILED', event.data.asset + ' (' + event.data.error + ')');
        }
    });

    function waiting() {
        update.hidden = !registration?.waiting;
        if (registration?.waiting) document.getElementById('update-note').textContent = 'Versi baru tersedia. Simpan hasil dan selesaikan dokumen sebelum memperbarui.';
    }
    // No forced activation/reload while a draft is open. Other tabs are checked in the worker too.
    function isWorking() {
        return window.KOMAD_DOCUMENT_OPEN === true;
    }
    navigator.serviceWorker?.addEventListener('message', event => {
        if (event.data?.type === 'CHECK_BUSY') event.ports[0]?.postMessage({
            busy: isWorking()
        });
    });
    update.onclick = () => {
        if (isWorking()) {
            document.getElementById('update-note').textContent = 'Simpan draft Word terlebih dahulu, lalu pilih Mulai dokumen baru. Setelah itu tekan Perbarui aplikasi.';
            return;
        }
        if (registration?.waiting) {
            reloadRequested = true;
            registration.waiting.postMessage({
                type: 'ACTIVATE_UPDATE'
            });
            document.getElementById('update-note').textContent = 'Memeriksa dokumen yang masih terbuka…';
        }
    };
    navigator.serviceWorker?.addEventListener('message', event => {
        if (event.data?.type === 'UPDATE_BUSY') {
            reloadRequested = false;
            document.getElementById('update-note').textContent = 'Pembaruan ditunda. Selesaikan dokumen pada semua jendela aplikasi terlebih dahulu.';
        }
    });
    navigator.serviceWorker?.addEventListener('controllerchange', () => {
        if (reloadRequested || (hadController && !isWorking())) location.reload();
        else verify();
        hadController = true;
    });
    async function verify() {
        try {
            const worker = registration?.active;
            if (!worker) return;
            const channel = new MessageChannel();
            const info = await new Promise((resolve, reject) => {
                const timer = setTimeout(() => reject(Error('timeout')), 6000);
                channel.port1.onmessage = e => {
                    clearTimeout(timer);
                    resolve(e.data);
                };
                worker.postMessage({
                    type: 'CHECK_OFFLINE'
                }, [channel.port2]);
            });
            ready = info.complete === true;
            cacheVersion = info.version || '';
            if (ready) {
                networkState();
                document.getElementById('app-version').textContent = 'Versi 1.3.2';
            } else {
                state('Belum siap offline', 'Sebagian komponen belum tersimpan. Sambungkan internet, lalu tekan Periksa pembaruan.', 'error');
                diagnostic('CACHE_INCOMPLETE');
            }
        } catch {
            state('Kesiapan offline belum terkonfirmasi', 'Pemeriksaan komponen tidak menjawab. Gunakan Periksa pembaruan saat tersambung internet.', 'error');
            diagnostic('CHECK_TIMEOUT');
        }
    }
    check.onclick = async () => {
        if (!registration) {
            if (isWorking()) {
                document.getElementById('update-note').textContent = 'Simpan draft lalu pilih Mulai dokumen baru sebelum memeriksa kembali.';
                return;
            }
            location.reload();
            return;
        }
        check.disabled = true;
        document.getElementById('update-note').textContent = 'Memeriksa versi aplikasi…';
        try {
            await registration.update();
            await verify();
            waiting();
            if (!registration.waiting) document.getElementById('update-note').textContent = 'Pemeriksaan selesai. Versi baru, jika tersedia, akan ditawarkan setelah selesai disimpan.';
        } catch {
            document.getElementById('update-note').textContent = 'Belum bisa memeriksa pembaruan. Sambungkan internet dan coba lagi.';
        } finally {
            check.disabled = false;
        }
    };
    const environment = window.KOMAD_OFFLINE_STATUS.environment({
        secure: window.isSecureContext === true,
        serviceWorker: 'serviceWorker' in navigator,
        protocol: location.protocol,
        embedded: window.self !== window.top
    });
    if (environment.code !== 'SUPPORTED') {
        state(environment.title, environment.detail, 'error');
        diagnostic(environment.code);
        check.textContent = 'Periksa kembali';
        return;
    }
    state('Menyiapkan penggunaan offline…', 'Tetap sambungkan internet sebentar. Tunggu sampai semua komponen aplikasi tersimpan.', 'loading');

    function watchInstalling(reg) {
        const worker = reg.installing;
        if (!worker) return;
        const changed = () => {
            if (worker.state === 'installed') {
                waiting();
                if (!reg.waiting) verify();
            }
            if (worker.state === 'redundant' && !ready) {
                state('Persiapan offline belum selesai', 'Ada komponen yang gagal disimpan. Sambungkan internet dan tekan Periksa pembaruan.', 'error');
                diagnostic('INSTALL_FAILED');
            }
        };
        worker.addEventListener('statechange', changed);
        changed();
    }
    navigator.serviceWorker.register('./sw.js', {
        scope: './',
        updateViaCache: 'none'
    }).then(async reg => {
        registration = reg;
        waiting();
        reg.addEventListener('updatefound', () => watchInstalling(reg));
        watchInstalling(reg);
        let timer;
        try {
            await Promise.race([navigator.serviceWorker.ready, new Promise((_, reject) => {
                timer = setTimeout(() => reject(Object.assign(Error('Persiapan terlalu lama'), {
                    name: 'PreparationTimeout'
                })), 25000);
            })]);
            await verify();
        } finally {
            clearTimeout(timer);
        }
    }).catch(error => {
        state('Belum siap digunakan offline', 'Persiapan belum berhasil. Gunakan Periksa kembali saat tersambung internet; info kendala tersedia di bawah.', 'error');
        diagnostic(lastPreparationFailure ? 'INSTALL_FAILED' : error.name === 'PreparationTimeout' ? 'PREPARATION_TIMEOUT' : 'REGISTRATION_FAILED', error.name);
        check.textContent = 'Periksa kembali';
    });
})();

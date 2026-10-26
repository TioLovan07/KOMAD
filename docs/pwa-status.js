'use strict';
// This checks browser capabilities, not whether an icon exists on the device.
window.KOMAD_OFFLINE_STATUS = {
    environment({
        secure,
        serviceWorker,
        protocol,
        embedded
    }) {
        if (!secure) return {
            code: 'INSECURE_CONTEXT',
            title: protocol === 'https:' ? 'Browser membatasi fitur offline' : 'Gunakan tautan HTTPS resmi',
            detail: protocol === 'https:' ? 'Alamat sudah HTTPS, tetapi browser tidak memberikan konteks aman. Buka tautan langsung melalui Chrome atau Edge, di luar pratinjau aplikasi lain.' : 'Buka tautan resmi berawalan https://. Membuat ikon dari halaman ini belum mengaktifkan penggunaan offline.'
        };
        if (!serviceWorker) return {
            code: 'API_UNAVAILABLE',
            title: 'Fitur offline tidak tersedia pada browser ini',
            detail: embedded ? 'Buka tautan langsung di Chrome atau Edge. Tampilan di dalam aplikasi lain dapat membatasi fitur offline.' : 'Browser tidak menyediakan fitur penyimpanan aplikasi offline. Coba Chrome atau Edge pada jendela biasa. Jika masih sama, salin info kendala di bawah untuk diperiksa.'
        };
        return {
            code: 'SUPPORTED'
        };
    }
};

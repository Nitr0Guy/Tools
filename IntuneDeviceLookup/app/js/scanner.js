/* ═══════════════════════════════════════════════════════
   scanner.js — html5-qrcode barcode scanner wrapper
   ═══════════════════════════════════════════════════════ */

const Scanner = (() => {
    let html5Qrcode = null;
    let running = false;
    let onScanCallback = null;

    const supportedFormats = [
        Html5QrcodeSupportedFormats.CODE_128,
        Html5QrcodeSupportedFormats.CODE_39,
        Html5QrcodeSupportedFormats.CODE_93,
        Html5QrcodeSupportedFormats.EAN_13,
        Html5QrcodeSupportedFormats.QR_CODE,
    ];

    function init(elementId, onScan) {
        onScanCallback = onScan;
        html5Qrcode = new Html5Qrcode(elementId, {
            formatsToSupport: supportedFormats,
            verbose: false,
        });
    }

    async function start() {
        if (!html5Qrcode || running) return;

        const config = {
            fps: 10,
            qrbox: { width: 280, height: 120 },
            rememberLastUsedCamera: true,
            aspectRatio: 1.5,
        };

        try {
            await html5Qrcode.start(
                { facingMode: 'environment' },
                config,
                (decodedText) => {
                    if (onScanCallback) {
                        // Vibrate on successful scan
                        if (navigator.vibrate) navigator.vibrate(100);
                        onScanCallback(decodedText.trim());
                    }
                },
                () => { /* ignore scan failures (each frame that doesn't decode) */ }
            );
            running = true;
        } catch (err) {
            console.error('Scanner start failed:', err);
            throw new Error('Could not start camera. Please allow camera access.');
        }
    }

    async function stop() {
        if (!html5Qrcode || !running) return;
        try {
            await html5Qrcode.stop();
        } catch (_) { /* ignore */ }
        running = false;
    }

    function isRunning() {
        return running;
    }

    return { init, start, stop, isRunning };
})();

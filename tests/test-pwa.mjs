import vm from 'node:vm';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
const dist = new URL('../docs/', import.meta.url),
    source = await fs.readFile(new URL('sw.js', dist), 'utf8');
const origin = 'https://komad.example.test/komad/',
    assets = JSON.parse(source.match(/ASSETS\s*=\s*(\[[\s\S]*?\]);/)[1]);
for (const asset of assets) assert((await fs.stat(new URL(asset, dist))).size > 0, asset);
const manifest = JSON.parse(await fs.readFile(new URL('manifest.webmanifest', dist), 'utf8'));
assert.equal(manifest.display, 'standalone');
assert.equal(manifest.scope, './');
for (const icon of manifest.icons) assert((await fs.stat(new URL(icon.src, dist))).size > 0);

function environment(failAt = '', badRedirect = false) {
    const handlers = {},
        data = new Map(),
        network = [],
        deleted = [];
    let skipped = false,
        claimed = false,
        clients = [];
    const caches = {
        open: async key => {
            if (!data.has(key)) data.set(key, new Map());
            const entries = data.get(key);
            return {
                put: async (u, r) => entries.set(String(u), r),
                match: async u => entries.get(String(u))
            };
        },
        keys: async () => [...data.keys()],
        delete: async k => {
            deleted.push(k);
            return data.delete(k);
        }
    };
    const fetch = async request => {
        const url = typeof request === 'string' ? request : request.url;
        network.push(url);
        if (url.endsWith(failAt) && failAt) throw Error('network failure');
        const relative = url.slice(origin.length) || 'index.html',
            body = await fs.readFile(new URL(relative, dist));
        const response = new Response(body, {
            headers: {
                'Content-Type': relative.endsWith('.js') ? 'application/javascript' : relative.endsWith('.html') ? 'text/html' : 'application/octet-stream'
            }
        });
        const finalURL = relative === 'index.html' ? origin : relative === 'viewer.html' ? origin + (badRedirect ? 'sign-in' : 'viewer') : url;
        Object.defineProperty(response, 'url', {
            value: finalURL
        });
        Object.defineProperty(response, 'redirected', {
            value: finalURL !== url
        });
        return response;
    };
    const self = {
        location: {
            href: origin + 'sw.js'
        },
        addEventListener: (name, fn) => handlers[name] = fn,
        clients: {
            claim: async () => {
                claimed = true;
            },
            matchAll: async () => clients
        },
        skipWaiting: async () => {
            skipped = true;
        }
    };
    vm.runInNewContext(source, {
        self,
        caches,
        fetch,
        URL,
        Request,
        Response,
        Headers,
        MessageChannel,
        setTimeout,
        clearTimeout
    });
    const dispatch = async (type, event = {}) => {
        let promise;
        handlers[type]({
            ...event,
            waitUntil: p => promise = p
        });
        await promise;
    };
    return {
        data,
        network,
        deleted,
        dispatch,
        handlers,
        setClients: c => clients = c,
        get skipped() {
            return skipped;
        },
        get claimed() {
            return claimed;
        }
    };
}
const good = environment();
await good.dispatch('install');
assert.equal(good.data.size, 1);
const entries = [...good.data.values()][0];
assert.equal(entries.size, assets.length);
assert(entries.has(origin), 'Canonical home URL must be cached directly');
assert(!good.network.includes(origin + 'index.html'), 'Do not fetch the redirecting index.html address');
assert.equal(entries.get(origin + 'viewer.html').redirected, false, 'Cached viewer must be usable for navigation after a canonical redirect');
await good.dispatch('activate');
assert(good.claimed);
let reply;
await good.dispatch('message', {
    data: {
        type: 'CHECK_OFFLINE'
    },
    ports: [{
        postMessage: x => reply = x
    }]
});
assert(reply.complete);
let responsePromise;
const previousRequests = good.network.length;
good.handlers.fetch({
    request: {
        url: origin + '?source=app',
        method: 'GET',
        mode: 'navigate'
    },
    respondWith: p => responsePromise = p
});
assert((await responsePromise).ok);
assert.equal(good.network.length, previousRequests, 'Offline navigation must use cached shell');
let intercepted = false;
good.handlers.fetch({
    request: {
        url: origin + 'user-document.docx',
        method: 'GET',
        mode: 'navigate'
    },
    respondWith: () => intercepted = true
});
assert(!intercepted, 'Documents must never enter shell cache');
const client = busy => ({
    url: origin,
    frameType: 'top-level',
    postMessage: (_, ports) => ports[0].postMessage({
        busy
    })
});
// A busy document in any top-level tab blocks activation. The preview iframe does not block it.
good.setClients([client(true)]);
let blocked = false;
await good.dispatch('message', {
    data: {
        type: 'ACTIVATE_UPDATE'
    },
    source: {
        postMessage: m => blocked = m.type === 'UPDATE_BUSY'
    }
});
assert(blocked);
assert(!good.skipped);
good.setClients([client(false), {
    url: origin + 'viewer.html',
    frameType: 'nested',
    postMessage: () => {
        throw Error('Nested frame must be excluded');
    }
}]);
await good.dispatch('message', {
    data: {
        type: 'ACTIVATE_UPDATE'
    },
    source: {
        postMessage: () => {}
    }
});
assert(good.skipped);
const broken = environment('viewer.js');
await assert.rejects(broken.dispatch('install'));
assert.equal(broken.data.size, 0, 'Partial install must not claim offline readiness');
const invalidRedirect = environment('', true);
await assert.rejects(invalidRedirect.dispatch('install'), /UNEXPECTED_REDIRECT/);
assert.equal(invalidRedirect.data.size, 0, 'A login or unrelated redirect must not be cached');
console.log('PASS: canonical home URL, HTML redirect normalization, unrelated redirect rejection.');
console.log('PASS: shell/manifest assets, cache-only launch, no document caching, complete-cache readiness, multi-tab update guard, iframe exclusion, atomic failure.');
const capabilityContext = {
    window: {}
};
vm.runInNewContext(await fs.readFile(new URL('pwa-status.js', dist), 'utf8'), capabilityContext);
const classify = capabilityContext.window.KOMAD_OFFLINE_STATUS.environment;
assert.equal(classify({
    secure: false,
    serviceWorker: true,
    protocol: 'http:',
    embedded: false
}).code, 'INSECURE_CONTEXT');
assert.equal(classify({
    secure: true,
    serviceWorker: false,
    protocol: 'https:',
    embedded: false
}).code, 'API_UNAVAILABLE');
assert.equal(classify({
    secure: true,
    serviceWorker: true,
    protocol: 'https:',
    embedded: false
}).code, 'SUPPORTED');
console.log('PASS: HTTPS without offline API stays unavailable; supported environments proceed to cache verification.');
process.exit(0);

const test = require('node:test');
const assert = require('node:assert/strict');

const { isApiRequest } = require('../src/middleware/request-logger');

test('identifica endpoints API sin incluir paginas ni archivos estaticos', () => {
    assert.equal(isApiRequest({ path: '/api/tipo-cambio/monex' }), true);
    assert.equal(isApiRequest({ path: '/auth/login' }), true);
    assert.equal(isApiRequest({ path: '/dashboard-stats' }), true);
    assert.equal(isApiRequest({ path: '/dashboard' }), false);
    assert.equal(isApiRequest({ path: '/styles/tailwind.css' }), false);
});

const { getAccessToken, getUserFromToken } = require('../services/auth');

async function requireAuth(req, res, next) {
    try {
        const accessToken = getAccessToken(req);
        const user = await getUserFromToken(accessToken);

        if (!user) {
            const wantsPage = req.method === 'GET' && (
                req.path === '/search' ||
                req.path === '/dashboard' ||
                req.path === '/gastos' ||
                req.path === '/tipo-cambio' ||
                req.path.endsWith('.html')
            );

            if (wantsPage) {
                return res.redirect('/login');
            }

            return res.status(401).json({ error: 'unauthorized' });
        }

        req.user = user;
        return next();
    } catch (error) {
        return next(error);
    }
}

module.exports = { requireAuth };

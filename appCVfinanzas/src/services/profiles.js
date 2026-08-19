const { hasSupabaseConfig, selectFromSupabase } = require('../lib/supabase');
const { hasPostgresConfig, queryPostgres, quoteTableName } = require('../lib/postgres');

const fallbackProfile = {
    name: 'El Ambicioso',
    label: 'Crecimiento Financiero',
    description: 'Tienes una mentalidad orientada a metas y buscas mejorar tu relacion con el dinero de forma constante.',
    strengths: [
        'Defines objetivos claros',
        'Buscas oportunidades para crecer',
        'Te motivan los avances medibles'
    ],
    challenges: [
        'Mantener un presupuesto consistente',
        'Equilibrar riesgo y estabilidad',
        'Celebrar avances pequenos sin perder foco'
    ]
};

const profilesByType = {
    G: {
        name: 'El Gastador',
        label: 'Compulsivo-Gastador',
        description: 'Disfrutas el dinero y la experiencia de gastar, pero tu reto es usar esa energia con mas intencion.',
        strengths: ['Espontaneo', 'Disfrutas el presente', 'Generoso'],
        challenges: ['Planificar antes de comprar', 'Evitar gastos impulsivos', 'Separar emocion de decision financiera']
    },
    A: {
        name: 'El Ahorrador',
        label: 'Compulsivo-Ahorrador',
        description: 'Tienes disciplina financiera, pero tu relacion con el dinero puede estar marcada por miedo a gastar.',
        strengths: ['Muy disciplinado', 'Excelente planificador', 'Bajo endeudamiento'],
        challenges: ['Disfrutar el dinero', 'Reducir ansiedad al gastar', 'No sacrificar siempre el presente']
    },
    D: {
        name: 'El Ambicioso',
        label: 'Compulsivo-Dinero',
        description: 'Tienes una mentalidad orientada a metas y buscas mejorar tu relacion con el dinero de forma constante.',
        strengths: ['Alta motivacion', 'Enfoque en resultados', 'Capaz de generar ingresos'],
        challenges: ['No medir tu valor solo por dinero', 'Cuidar relaciones y descanso', 'Celebrar avances sin mover siempre la meta']
    },
    I: {
        name: 'El Indiferente',
        label: 'Indiferente al Dinero',
        description: 'Prefieres no lidiar demasiado con el dinero, pero ordenar tus decisiones puede darte mas libertad.',
        strengths: ['Sin obsesion por el dinero', 'Prioriza otras areas de la vida', 'Poco materialista'],
        challenges: ['Tomar decisiones financieras', 'Ordenar gastos e ingresos', 'Dar direccion al dinero']
    },
    M: {
        name: 'El Idealista',
        label: 'Monje del Dinero',
        description: 'Tienes valores claros y puedes aprender a ver el dinero como una herramienta compatible con ellos.',
        strengths: ['Valores claros', 'No materialista', 'Coherente con sus principios'],
        challenges: ['Trabajar bloqueos para generar dinero', 'No asociar riqueza con maldad', 'Permitir abundancia con proposito']
    },
    F: {
        name: 'El Fluido',
        label: 'Amante del Dinero',
        description: 'Tienes una relacion sana con el dinero: lo administras con intencion y lo compartes sin miedo.',
        strengths: ['Relacion sana con el dinero', 'Generoso y agradecido', 'Toma decisiones con calma'],
        challenges: ['Mantener equilibrio bajo presion', 'No contagiarse del estres ajeno', 'Seguir creciendo sin perder el piso']
    }
};

function profileFromQuizResult(row) {
    const baseProfile = profilesByType[row.profile_type] || fallbackProfile;

    return {
        ...baseProfile,
        profileType: row.profile_type,
        label: row.profile_label || baseProfile.label,
        color: row.profile_color,
        score: row.profile_score,
        scores: row.scores,
        email: row.user_email,
        createdAt: row.created_at
    };
}

function getFinancialProfile(slug) {
    // Mapear slugs a tipos de perfil
    const slugMap = {
        'el-gastador': 'G',
        'el-ahorrador': 'A',
        'el-ambicioso': 'D',
        'el-indiferente': 'I',
        'el-idealista': 'M',
        'el-fluido': 'F'
    };

    const profileType = slugMap[slug?.toLowerCase()] || 'D';
    return profilesByType[profileType] || fallbackProfile;
}

function normalizeProfileType(value) {
    if (!value) {
        return '';
    }

    const normalized = value
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/^el\s+/, '')
        .replace(/\s+/g, '-');

    const typeMap = {
        g: 'G',
        gastador: 'G',
        'el-gastador': 'G',
        'compulsivo-gastador': 'G',
        a: 'A',
        ahorrador: 'A',
        'el-ahorrador': 'A',
        'compulsivo-ahorrador': 'A',
        d: 'D',
        ambicioso: 'D',
        dinero: 'D',
        'el-ambicioso': 'D',
        'compulsivo-dinero': 'D',
        i: 'I',
        indiferente: 'I',
        'el-indiferente': 'I',
        'indiferente-al-dinero': 'I',
        m: 'M',
        idealista: 'M',
        monje: 'M',
        'el-idealista': 'M',
        'monje-del-dinero': 'M',
        f: 'F',
        fluido: 'F',
        amante: 'F',
        'el-fluido': 'F',
        'amante-del-dinero': 'F'
    };

    return typeMap[normalized] || value.trim().toUpperCase();
}

async function getFinancialProfileByEmail(email) {
    if (hasPostgresConfig) {
        const results = await searchResults({ email });
        return results[0] || fallbackProfile;
    }

    if (!hasSupabaseConfig) {
        return fallbackProfile;
    }

    const tableName = process.env.SUPABASE_RESULTS_TABLE || 'quiz_honda_results';
    const params = new URLSearchParams({
        select: 'user_email,profile_type,profile_label,profile_color,profile_score,scores,created_at',
        user_email: `eq.${email}`,
        order: 'created_at.desc',
        limit: '1'
    });
    const data = await selectFromSupabase(tableName, params);

    return data && data[0] ? profileFromQuizResult(data[0]) : fallbackProfile;
}

async function searchResultsFromPostgres({ email, profileType } = {}) {
    const normalizedEmail = email?.trim();
    const normalizedProfileType = normalizeProfileType(profileType);
    const tableName = process.env.AZURE_PG_RESULTS_TABLE ||
        process.env.PG_RESULTS_TABLE ||
        process.env.SUPABASE_RESULTS_TABLE ||
        'quiz_honda_results';
    const where = [];
    const values = [];

    if (normalizedEmail) {
        values.push(normalizedEmail);
        where.push(`user_email = $${values.length}`);
    }

    if (normalizedProfileType) {
        values.push(normalizedProfileType);
        where.push(`profile_type = $${values.length}`);
    }

    const sql = `
        select
            user_email,
            profile_type,
            profile_label,
            profile_color,
            profile_score,
            scores,
            created_at
        from ${quoteTableName(tableName)}
        ${where.length ? `where ${where.join(' and ')}` : ''}
        order by created_at desc
    `;

    const data = await queryPostgres(sql, values);
    return Array.isArray(data) ? data.map(profileFromQuizResult) : [];
}

async function searchResultsFromSupabase({ email, profileType } = {}) {
    const normalizedEmail = email?.trim();
    const normalizedProfileType = normalizeProfileType(profileType);
    const tableName = process.env.SUPABASE_RESULTS_TABLE || 'quiz_honda_results';
    const params = new URLSearchParams({
        select: 'user_email,profile_type,profile_label,profile_color,profile_score,scores,created_at',
        order: 'created_at.desc'
    });

    if (normalizedEmail) {
        params.set('user_email', `eq.${normalizedEmail}`);
    }

    if (normalizedProfileType) {
        params.set('profile_type', `eq.${normalizedProfileType}`);
    }

    const data = await selectFromSupabase(tableName, params);
    return Array.isArray(data) ? data.map(profileFromQuizResult) : [];
}

async function searchResults({ email, profileType } = {}) {
    if (hasPostgresConfig) {
        return searchResultsFromPostgres({ email, profileType });
    }

    if (hasSupabaseConfig) {
        return searchResultsFromSupabase({ email, profileType });
    }

    return [];
}

async function getDashboardStats() {
    if (!hasPostgresConfig) {
        const results = await searchResults();

        return {
            totalResults: results.length,
            recentResults: results.filter((row) => {
                const createdAt = row.createdAt ? new Date(row.createdAt).getTime() : 0;
                return createdAt >= Date.now() - (7 * 24 * 60 * 60 * 1000);
            }).length,
            byProfile: Object.values(results.reduce((profiles, row) => {
                const key = row.profileType || 'N/A';
                profiles[key] = profiles[key] || {
                    profileType: key,
                    profileName: row.name || key,
                    total: 0
                };
                profiles[key].total += 1;
                return profiles;
            }, {})),
            latestResults: results.slice(0, 5)
        };
    }

    const tableName = process.env.AZURE_PG_RESULTS_TABLE ||
        process.env.PG_RESULTS_TABLE ||
        process.env.SUPABASE_RESULTS_TABLE ||
        'quiz_honda_results';
    const quotedTableName = quoteTableName(tableName);
    const [totals, byProfile, latestResults] = await Promise.all([
        queryPostgres(`
            select
                count(*)::int as total_results,
                count(*) filter (where created_at >= now() - interval '7 days')::int as recent_results
            from ${quotedTableName}
        `),
        queryPostgres(`
            select profile_type, profile_label, count(*)::int as total
            from ${quotedTableName}
            group by profile_type, profile_label
            order by total desc, profile_type asc
        `),
        queryPostgres(`
            select
                user_email,
                profile_type,
                profile_label,
                profile_color,
                profile_score,
                scores,
                created_at
            from ${quotedTableName}
            order by created_at desc
            limit 5
        `)
    ]);

    return {
        totalResults: totals[0]?.total_results || 0,
        recentResults: totals[0]?.recent_results || 0,
        byProfile: byProfile.map((row) => ({
            profileType: row.profile_type,
            profileName: profilesByType[row.profile_type]?.name || row.profile_label || row.profile_type,
            profileLabel: row.profile_label,
            total: row.total
        })),
        latestResults: latestResults.map(profileFromQuizResult)
    };
}

async function searchResultsByEmail(email) {
    return searchResults({ email });
}

module.exports = {
    getDashboardStats,
    getFinancialProfile,
    getFinancialProfileByEmail,
    normalizeProfileType,
    searchResults,
    searchResultsFromPostgres,
    searchResultsFromSupabase,
    searchResultsByEmail
};

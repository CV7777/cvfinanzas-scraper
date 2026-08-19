const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_ANON_KEY;

const hasSupabaseConfig = Boolean(supabaseUrl && supabaseKey);

async function selectFromSupabase(tableName, params) {
    if (!hasSupabaseConfig) {
        return null;
    }

    const url = `${supabaseUrl}/rest/v1/${tableName}?${params.toString()}`;
    const response = await fetch(url, {
        headers: {
            apikey: supabaseKey,
            Authorization: `Bearer ${supabaseKey}`
        }
    });

    if (!response.ok) {
        throw new Error(`Supabase request failed: ${response.status} ${await response.text()}`);
    }

    return response.json();
}

module.exports = {
    hasSupabaseConfig,
    supabaseKey,
    supabaseUrl,
    selectFromSupabase
};

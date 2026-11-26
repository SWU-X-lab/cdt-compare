(function (global) {
    'use strict';

    const SUPABASE_URL = 'https://rnfwxphdzvonoiduidsg.supabase.co';
    const SUPABASE_KEY = 'sb_publishable_kdXRnhVgTwsfgrrzO0ok8A_PZOKfHZK';
    global.CLTStorageFactories = global.CLTStorageFactories || {};
    global.CLTStorageFactories.supabase = function createSupabaseStorage() {
        let client = null;
        return {
            async initialize() {
                if (SUPABASE_URL === 'YOUR_SUPABASE_URL_HERE') return;
                if (!global.supabase || typeof global.supabase.createClient !== 'function') {
                    throw new Error('Supabase JavaScript client is unavailable');
                }
                client = global.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: false } });
            },
            saveTrial() { return Promise.resolve(); },
            saveCheckpoint() { return Promise.resolve(); },
            loadProgress() { return Promise.resolve(null); },
            saveProgress() { return Promise.resolve(false); },
            clearProgress() { return Promise.resolve(false); },
            async saveFinalResult(payload) {
                const rows = payload && Array.isArray(payload.trials) ? payload.trials : [];
                if (!client) return { saved: false, message: '本版本未配置服务器上传，仅生成本地 Excel。' };
                if (rows.length === 0) return { saved: false, message: '没有正式实验数据需要上传。' };
                const { error } = await client.from('vwm_data').insert(rows);
                // 不能删列重试：点击方位、落点和轨迹长度是本任务的研究数据。
                if (error) throw new Error(error.message || '完整数据写入失败；请核对服务器字段。');
                return { saved: true, message: '完整试次数据已上传服务器。' };
            },
            finishExperiment() { return Promise.resolve({ message: '' }); }
        };
    };
})(window);

/* Supabase back-end: Microsoft 365 sign-in (@hitap.net only) and per-user sync of documents and profile.
   Row Level Security limits every row to its owner, so the publishable key is safe in the browser.
   Signed out, the app keeps working from localStorage only. */
(() => {
  'use strict';
  const HD = window.HD;
  const CFG = window.HD_SUPABASE || {};
  const SDK = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.57.4/dist/umd/supabase.min.js';
  const DOMAIN = CFG.domain || 'hitap.net';
  const allowed = u => !!u && (u.email || '').toLowerCase().endsWith('@' + DOMAIN);
  const listeners = new Set();
  let client = null, user = null, status = 'off', pushTimer = null;
  const pending = new Map();
  let profilePending = null;

  const emit = () => listeners.forEach(fn => { try { fn({ user, status }); } catch (e) { console.error(e); } });
  const setStatus = s => { status = s; emit(); };

  async function init() {
    if (!CFG.url || !CFG.key) { HD.appBridge && HD.appBridge.remoteFailed(); return; }
    try { await HD.util.loadScript(SDK); } catch (e) { setStatus('offline'); HD.appBridge && HD.appBridge.remoteFailed(); return; }
    client = window.supabase.createClient(CFG.url, CFG.key, {
      // PKCE puts the sign-in result in ?code= so it doesn't collide with the app's #/ routes
      auth: { flowType: 'pkce', detectSessionInUrl: true, persistSession: true, autoRefreshToken: true },
    });
    client.auth.onAuthStateChange((event, session) => {
      if (session && !allowed(session.user)) { rejectForeign(); return; }
      const next = session ? session.user : null;
      const changed = (next && next.id) !== (user && user.id);
      user = next;
      if (location.search.includes('code=')) history.replaceState(null, '', location.pathname + location.hash);
      if (changed) checkAdmin();
      if (changed && user) pull();
      else emit();
    });
    loadTemplates();
    const q = new URLSearchParams(location.search + '&' + location.hash.replace(/^#/, ''));
    if (q.get('error_description')) {
      const msg = q.get('error_description');
      history.replaceState(null, '', location.pathname);
      alert(/database error/i.test(msg) ? `ใช้ได้เฉพาะบัญชี @${DOMAIN} เท่านั้น\nOnly @${DOMAIN} accounts can sign in.` : `เข้าสู่ระบบไม่สำเร็จ / Sign-in failed: ${msg}`);
    }
    const { data } = await client.auth.getSession();
    user = data.session && allowed(data.session.user) ? data.session.user : null;
    if (data.session && !user) rejectForeign();
    checkAdmin();
    if (user) pull(); else setStatus('signed-out');
  }

  function rejectForeign() {
    user = null;
    setTimeout(() => client.auth.signOut(), 0);
    setStatus('signed-out');
    alert(`ใช้ได้เฉพาะบัญชี @${DOMAIN} เท่านั้น\nOnly @${DOMAIN} accounts can sign in.`);
  }

  /** Merge cloud rows with the local copies (newest wins) and upload local-only/newer ones. */
  async function pull() {
    if (!client || !user) return;
    setStatus('syncing');
    try {
      const [{ data: rows, error }, { data: prof, error: perr }] = await Promise.all([
        client.from('documents').select('id, data, updated_at').order('updated_at', { ascending: false }).limit(200),
        client.from('profiles').select('data').maybeSingle(),
      ]);
      if (error) throw error;
      if (perr) throw perr;
      const bridge = HD.appBridge;
      const local = bridge.getDocs(), byId = new Map(local.map(d => [d.id, d]));
      const merged = [...local];
      for (const r of rows || []) {
        const remote = r.data;
        const mine = byId.get(r.id);
        if (!mine) merged.push(remote);
        else if ((remote.updated || 0) > (mine.updated || 0)) merged[merged.indexOf(mine)] = remote;
      }
      const remoteIds = new Map((rows || []).map(r => [r.id, r.data.updated || 0]));
      for (const d of local) if (!remoteIds.has(d.id) || (d.updated || 0) > remoteIds.get(d.id)) pending.set(d.id, d);
      const localProfile = bridge.getProfile();
      const profile = { ...(prof ? prof.data : {}), ...localProfile };
      bridge.replaceAll(merged, profile);
      if (!prof || Object.keys(localProfile).some(k => (prof.data || {})[k] !== localProfile[k])) profilePending = profile;
      await flush();
      setStatus('synced');
    } catch (e) {
      console.error('cloud pull', e);
      setStatus('error');
    }
  }

  async function flush() {
    if (!client || !user) return;
    const docs = [...pending.values()];
    pending.clear();
    try {
      if (docs.length) {
        const { error } = await client.from('documents').upsert(docs.map(d => ({
          user_id: user.id, id: d.id, tpl: d.tpl, title: d.title || null, data: d, updated_at: new Date(d.updated || Date.now()).toISOString(),
        })), { onConflict: 'user_id,id' });
        if (error) throw error;
      }
      if (profilePending) {
        const p = profilePending; profilePending = null;
        const { error } = await client.from('profiles').upsert({ user_id: user.id, data: p, updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
        if (error) throw error;
      }
      if (status !== 'synced') setStatus('synced');
    } catch (e) {
      console.error('cloud push', e);
      docs.forEach(d => pending.set(d.id, d));
      setStatus('error');
    }
  }
  const schedule = () => { clearTimeout(pushTimer); pushTimer = setTimeout(flush, 1200); };

  /** Template catalog rows (public read): admin uploads and hidden built-ins. */
  const publicUrl = path => `${CFG.url}/storage/v1/object/public/templates/${path.split('/').map(encodeURIComponent).join('/')}`;
  async function loadTemplates() {
    try {
      const { data, error } = await client.from('templates').select('id, kind, builtin, hidden, meta, data_path, original_path, updated_at').order('updated_at');
      if (error) throw error;
      HD.appBridge.setRemoteTemplates(data || [], publicUrl);
    } catch (e) {
      console.error('templates', e);
      HD.appBridge.remoteFailed();
    }
  }
  let admin = false;
  async function checkAdmin() {
    admin = false;
    if (user) { try { const { data } = await client.rpc('is_admin'); admin = data === true; } catch (e) { admin = false; } }
    HD.appBridge.setAdmin(admin);
    emit();
  }

  HD.cloud = {
    client: () => client,
    isAdmin: () => admin,
    publicUrl,
    reloadTemplates: () => loadTemplates(),
    enabled: () => !!(CFG.url && CFG.key),
    user: () => user,
    status: () => status,
    onChange(fn) { listeners.add(fn); fn({ user, status }); return () => listeners.delete(fn); },
    saveDoc(d) { if (!user || !d) return; pending.set(d.id, d); setStatus('saving'); schedule(); },
    saveProfile(p) { if (!user) return; profilePending = p; schedule(); },
    /** Microsoft 365 (Entra ID) via Supabase's "azure" provider; only @hitap.net is accepted. */
    async signInMicrosoft() {
      if (!client) throw new Error('Supabase is not available');
      // avoid landing on Supabase's raw JSON error page while the provider is still switched off
      const res = await fetch(`${CFG.url}/auth/v1/settings`, { headers: { apikey: CFG.key } }).then(r => r.json()).catch(() => null);
      if (res && res.external && !res.external.azure) throw new Error('provider is not enabled');
      const { error } = await client.auth.signInWithOAuth({
        provider: 'azure',
        options: { scopes: 'email openid profile', redirectTo: location.origin + location.pathname, queryParams: { domain_hint: DOMAIN, prompt: 'select_account' } },
      });
      if (error) throw error;
    },
    /** Admin page: password sign-in to the configured admin account (no email round-trip). */
    async signInAdmin(password) {
      if (!client) throw new Error('Supabase is not available');
      const { error } = await client.auth.signInWithPassword({ email: CFG.adminEmail, password });
      if (error) throw error;
    },
    async signOut() {
      if (!client) return;
      await flush();
      await client.auth.signOut();
      user = null;
      HD.appBridge.clearLocal();
      setStatus('signed-out');
    },
    sync: pull,
  };
  if (HD.onCloudReady) HD.onCloudReady();
  window.addEventListener('beforeunload', () => { if (pending.size || profilePending) flush(); });
  init();
})();

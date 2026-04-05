const API_URL = "https://script.google.com/macros/s/AKfycbzqoiq6p6A-_qJOGLoq9AUTC266Xz85sia9LQ9THoRKP-vlOI-v04TJd3XL5VFUtnCoIQ/exec";

let appData           = [];
let specData          = [];
let knowledgeProposed = [];
let editingState = { catIndex: null, tplIndex: null };
let lastSyncTime = 0;
const FAV_KEY     = 'bh_cs_favorites';
const USER_KEY    = 'bh_cs_username';
const HISTORY_KEY = 'bh_cs_history';
const CACHE_KEY   = 'bh_cs_cache';
const DRAFT_KEY   = 'bh_cs_draft_conflict';

const debouncedRender = (() => {
    let t = null;
    return () => { clearTimeout(t); t = setTimeout(render, 300); };
})();

let saveTimeout     = null;
let isSaving        = false;
let pendingSave     = false;
let commentTarget   = { c: null, t: null };
let confirmCallback = null;
let inputCallback   = null;

// =============================================
// DATETIME CONVERSION
// =============================================
function convertToJpDatetime(val) {
    const months = { Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12 };
    const m = val.trim().match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})\s+(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (m) {
        const day = String(m[1]).padStart(2,'0');
        const mon = String(months[m[2].charAt(0).toUpperCase()+m[2].slice(1).toLowerCase()]||'').padStart(2,'0');
        if (mon) return `${m[3]}/${mon}/${day} ${String(m[4]).padStart(2,'0')}:${m[5]}`;
    }
    return null;
}

// =============================================
// USER NAME
// =============================================
function initEditorName() {
    let name = localStorage.getItem(USER_KEY);
    if (!name) {
        name = "担当者";
        localStorage.setItem(USER_KEY, name);
        setTimeout(() => showInputModal(
            'ユーザー名設定',
            '履歴に残すあなたの名前を入力してください',
            '',
            n => { localStorage.setItem(USER_KEY, n); document.getElementById('display-username').innerText = n; }
        ), 600);
    }
    document.getElementById('display-username').innerText = name;
    return name;
}
function changeEditorName() {
    showInputModal('名前の変更', '新しい名前を入力してください', localStorage.getItem(USER_KEY) || '担当者', n => {
        localStorage.setItem(USER_KEY, n);
        document.getElementById('display-username').innerText = n;
        showToast('名前を変更しました');
    });
}
function getFavorites() { return JSON.parse(localStorage.getItem(FAV_KEY) || "[]"); }

// =============================================
// INIT
// =============================================
function ensureIds() {
    appData.forEach(cat => {
        if (!cat.id) cat.id = 'cat_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
        (cat.templates || []).forEach(tpl => {
            if (!tpl.id) tpl.id = 't_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
        });
    });
}

async function init() {
    initEditorName();
    const sb = document.getElementById('status-bar');
    if (!API_URL.startsWith("https://script.google.com")) {
        sb.className = 'status-badge offline';
        sb.innerHTML = '<span class="sync-dot"></span> ⚠️ オフラインモード';
        render(); return;
    }

    sb.className = 'status-badge syncing';
    sb.innerHTML = '<span class="sync-dot pulse"></span> 🔄 同期中...';

    try {
        // レスポンスをまずテキストで受け取り、原因を特定しやすくする
        const res     = await fetch(API_URL);
        const rawText = await res.text();

        let resJson;
        try {
            resJson = JSON.parse(rawText);
        } catch (_) {
            // GASがHTMLを返している → デプロイ未完了・権限エラーなど
            console.error('[BH-CS] GASレスポンスがJSONではありません:', rawText.slice(0, 300));
            throw new Error('GAS_NOT_JSON');
        }

        if (resJson.error) {
            console.error('[BH-CS] GASエラー:', resJson.error);
            throw new Error('GAS_ERROR:' + resJson.error);
        }

        if (resJson.json) {
            appData = JSON.parse(resJson.json);
            ensureIds();
            lastSyncTime = resJson.updatedAt;
            localStorage.setItem(CACHE_KEY, JSON.stringify(appData));
            sb.className = 'status-badge online';
            sb.innerHTML = '<span class="sync-dot"></span> ✅ サーバー同期済み';
        }
        render();

    } catch (e) {
        console.error('[BH-CS] 同期失敗:', e.message);

        // エラー種別に応じたメッセージ
        let hint = '';
        if (e.message === 'GAS_NOT_JSON') {
            hint = 'GASが正しく応答しません。デプロイ設定・承認を確認してください。';
        } else if (e.message.startsWith('GAS_ERROR:')) {
            hint = 'GASエラー: ' + e.message.slice(10);
        } else if (e.message.includes('Failed to fetch') || e.message.includes('NetworkError')) {
            hint = 'ネットワーク接続を確認してください。';
        } else {
            hint = e.message;
        }

        const cached = localStorage.getItem(CACHE_KEY);
        if (cached) {
            try { appData = JSON.parse(cached); ensureIds(); } catch(_) {}
            sb.className = 'status-badge offline';
            sb.innerHTML = `<span class="sync-dot"></span> 📦 キャッシュ表示`;
            sb.title = hint; // ホバーで詳細確認できる
            showToast('⚠️ ' + hint);
        } else {
            sb.className = 'status-badge offline';
            sb.innerHTML = `<span class="sync-dot"></span> ❌ 通信失敗`;
            sb.title = hint;
            showToast('❌ ' + hint);
        }
        render();
    }
}

// =============================================
// SAVE
// =============================================
async function executeSave(force = false, silent = false) {
    if (!API_URL.startsWith("https://script.google.com")) return;
    isSaving = true;
    try {
        const res = await fetch(API_URL, { method:"POST", body: JSON.stringify({ lastSyncTime: force ? 0 : lastSyncTime, data: appData }) });
        const result = await res.text();
        if (result === "conflict") {
            pendingSave = false;
            localStorage.setItem(DRAFT_KEY, JSON.stringify({ data: appData, savedAt: new Date().toLocaleString('ja-JP') }));
            openConflictModal();
        }
        else { lastSyncTime = parseInt(result); if (!silent) showToast('クラウドに保存しました'); }
    } catch(e) { if (!silent) console.error("通信エラー",e); }
    finally { isSaving = false; if (pendingSave) { pendingSave = false; executeSave(false, true); } }
}
function saveData(force=false, silent=false, instant=false) {
    if (!API_URL.startsWith("https://script.google.com")) { if (!silent) showToast('オフラインのため保存できません'); return; }
    clearTimeout(saveTimeout);
    if (instant) { if (isSaving) pendingSave = true; else executeSave(force, silent); }
    else { saveTimeout = setTimeout(() => { if (isSaving) pendingSave = true; else executeSave(false, silent); }, 2000); }
}

// =============================================
// KPI
// =============================================
function updateKPIs() {
    const all    = appData.flatMap(c => c.templates || []);
    const total  = all.length;
    const copies = all.reduce((s,t) => s+(t.copyCount||0), 0);
    const top    = [...all].sort((a,b) => (b.copyCount||0)-(a.copyCount||0))[0];
    document.getElementById('kpi-total').innerText  = total;
    document.getElementById('kpi-copies').innerText = copies;
    document.getElementById('kpi-cats').innerText   = appData.length;
    document.getElementById('kpi-top').innerText    = (top && top.copyCount > 0) ? (top.title || '—') : '—';
}

// =============================================
// ESCAPE HTML
// =============================================
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// =============================================
// HIGHLIGHT
// =============================================
function highlight(escapedText, terms) {
    if (!terms || terms.length === 0) return escapedText;
    let result = escapedText;
    terms.forEach(term => {
        const safe = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        result = result.replace(new RegExp(safe, 'gi'), m =>
            `<mark style="background:rgba(56,189,248,0.28);color:inherit;border-radius:2px;padding:0 1px;">${m}</mark>`
        );
    });
    return result;
}

// =============================================
// RENDER
// =============================================
function render() {
    const appDiv       = document.getElementById('app');
    const tocContainer = document.getElementById('toc-container');
    appDiv.innerHTML = ''; tocContainer.innerHTML = '';

    const searchQuery  = document.getElementById('searchInput').value.toLowerCase();
    const showOnlyFavs = document.getElementById('favFilter').checked;
    const favorites    = getFavorites();
    const isFiltering  = searchQuery !== "" || showOnlyFavs;
    const searchTerms  = searchQuery.split(/\s+/).filter(t => t);
    const visibleCats  = [];
    const allMax       = Math.max(...appData.flatMap(c=>c.templates).map(t=>t.copyCount||0), 1);

    appData.forEach((cat, cIndex) => {
        const filteredTpls = cat.templates.filter(tpl => {
            if (showOnlyFavs && !favorites.includes(tpl.id)) return false;
            const matchStr = (tpl.title+" "+tpl.content+" "+(tpl.tags||'')).toLowerCase();
            return searchTerms.every(term => matchStr.includes(term));
        });
        if (isFiltering && filteredTpls.length === 0 && editingState.catIndex !== cIndex) return;
        visibleCats.push({ id: cat.id||`cat-${cIndex}`, title: cat.title });

        const catSection = document.createElement('div');
        catSection.className = 'cat-section';
        catSection.id = cat.id||`cat-${cIndex}`;

        const totalCopies = cat.templates.reduce((s,t)=>s+(t.copyCount||0),0);
        catSection.innerHTML = `
            <div class="cat-header">
                <div class="cat-header-title">
                    <span>${esc(cat.title)}</span>
                    <span class="cat-count">${cat.templates.length}件</span>
                </div>
                <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
                    <span style="font-size:0.72rem; color:var(--dim);">コピー計 <b style="color:var(--success);">${totalCopies}</b>回</span>
                    <button class="btn btn-ghost btn-xs" onclick="editCategory(${cIndex})">✏️ 名前変更</button>
                    <button class="btn btn-danger btn-xs" onclick="deleteCategory(${cIndex})">🗑️ 削除</button>
                </div>
            </div>`;

        const table = document.createElement('table');
        table.className = 'tpl-table';
        const tbody = document.createElement('tbody');
        tbody.dataset.catIndex = cIndex;

        cat.templates.forEach((tpl, tIndex) => {
            if (isFiltering && !filteredTpls.includes(tpl) && !(editingState.catIndex===cIndex && editingState.tplIndex===tIndex)) return;
            const tr        = document.createElement('tr');
            const isEditing = (editingState.catIndex===cIndex && editingState.tplIndex===tIndex);
            const isFav     = favorites.includes(tpl.id);
            const copyCount = tpl.copyCount || 0;
            const barWidth  = Math.min(100, Math.round((copyCount/allMax)*100));

            if (isEditing) {
                tr.innerHTML = `
                    <th>
                        <label class="edit-label">件名</label>
                        <input type="text" id="edit-title-${cIndex}-${tIndex}" class="edit-input" value="${esc(tpl.title)}" placeholder="件名">
                        <label class="edit-label">タグ（,区切り）</label>
                        <div style="display:flex; gap:6px; margin-bottom:8px;">
                            <input type="text" id="edit-tags-${cIndex}-${tIndex}" class="edit-input" value="${esc(tpl.tags||'')}" placeholder="タグ（,区切り）" style="margin-bottom:0;">
                            <button class="btn btn-purple btn-sm" style="flex-shrink:0;" onclick="autoTag(${cIndex},${tIndex})" id="btn-auto-tag-${cIndex}-${tIndex}">🏷️ AI</button>
                        </div>
                    </th>
                    <td>
                        <label class="edit-label" style="display:flex;justify-content:space-between;">
                            <span>本文 <span style="color:var(--dim);">（変数は {{変数名}} 形式）</span></span>
                            <span id="char-count-${cIndex}-${tIndex}" style="font-weight:400;color:var(--dim);">${tpl.content.length} 文字</span>
                        </label>
                        <textarea id="edit-content-${cIndex}-${tIndex}" class="edit-textarea"
                            oninput="document.getElementById('char-count-${cIndex}-${tIndex}').innerText=this.value.length+' 文字'">${esc(tpl.content)}</textarea>
                        <label class="edit-label">メモ・注意書き</label>
                        <textarea id="edit-comment-${cIndex}-${tIndex}" class="edit-input" style="height:60px; resize:vertical;" placeholder="使用上の注意、ヒントなど...">${esc(tpl.comment||'')}</textarea>
                        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:4px;">
                            <button class="btn btn-primary" onclick="saveTemplate(${cIndex},${tIndex})">💾 保存</button>
                            <button class="btn btn-ghost" onclick="cancelEdit()">❌ 中止</button>
                        </div>
                    </td>`;
            } else {
                const tagHtml     = tpl.tags ? tpl.tags.split(',').map(tag=>`<span class="tag-badge">🏷️ ${esc(tag.trim())}</span>`).join('') : '';
                const logHtml     = tpl.lastEditedBy ? `<span class="edit-log">最終更新: ${esc(tpl.lastEditedAt)} (${esc(tpl.lastEditedBy)})</span>` : '';
                const commentHtml = tpl.comment ? `<div class="comment-area">💬 ${esc(tpl.comment)}</div>` : '';

                tr.innerHTML = `
                    <th>
                        <div style="padding-right:30px;">
                            <div style="font-size:0.88rem; font-weight:700; line-height:1.4;">${esc(tpl.title).replace(/\n/g,'<br>')}</div>
                            ${tagHtml ? `<div style="margin-top:6px;">${tagHtml}</div>` : ''}
                        </div>
                        <button class="btn-fav ${isFav?'active':''}" onclick="toggleFavorite('${tpl.id}')">★</button>
                        <div class="usage-bar-wrap"><div class="usage-bar" style="width:${barWidth}%"></div></div>
                        <div style="font-size:0.68rem; color:var(--dim); margin-top:3px;">使用回数: <b style="color:var(--success);">${copyCount}</b>回</div>
                    </th>
                    <td>
                        <div class="template-box">${searchTerms.length > 0 ? highlight(esc(tpl.content), searchTerms) : esc(tpl.content)}</div>
                        ${commentHtml}
                        <div style="display:flex; flex-wrap:wrap; gap:5px; margin-top:8px;">
                            <button class="btn btn-success" onclick="copyTextWithVariables(${cIndex},${tIndex})">📋 コピー</button>
                            <button class="btn btn-warn" onclick="startEdit(${cIndex},${tIndex})">✏️ 編集</button>
                            <button class="btn btn-ghost btn-sm" onclick="openCommentModal(${cIndex},${tIndex})">💬 メモ</button>
                            <button class="btn btn-danger btn-sm" onclick="deleteTemplate(${cIndex},${tIndex})">🗑️</button>
                        </div>
                        <div class="meta-row">
                            <span class="copy-count-badge">🔄 ${copyCount}回</span>
                            ${logHtml}
                        </div>
                    </td>`;
            }
            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        catSection.appendChild(table);

        if (!isFiltering) {
            const addBtn = document.createElement('button');
            addBtn.className = 'btn btn-purple btn-block';
            addBtn.innerHTML = '➕ 項目追加';
            addBtn.onclick = () => addNewTemplate(cIndex);
            catSection.appendChild(addBtn);
        }
        appDiv.appendChild(catSection);
    });

    // TOC
    if (visibleCats.length > 0) {
        const tocDiv = document.createElement('div');
        tocDiv.className = 'toc-card';
        tocDiv.innerHTML = `<div class="toc-title">📑 目次 <span style="font-weight:400; font-size:0.95em;">(ドラッグで並び替え)</span></div>`;
        const ul = document.createElement('ul');
        ul.className = 'toc-list'; ul.id = 'toc-list';
        visibleCats.forEach(item => {
            const li = document.createElement('li');
            li.onclick = () => { window.location.hash = item.id; };
            li.innerHTML = `<a href="#${item.id}">${esc(item.title)}</a>`;
            ul.appendChild(li);
        });
        tocDiv.appendChild(ul);
        tocContainer.appendChild(tocDiv);
    }

    if (!isFiltering) {
        const addCatBtn = document.createElement('button');
        addCatBtn.className = 'btn btn-ghost';
        addCatBtn.style.cssText = 'display:flex; width:100%; padding:12px; border-radius:var(--radius-lg); margin-top:8px; font-size:0.9rem; justify-content:center;';
        addCatBtn.innerHTML = '📂 新しいカテゴリーを追加する';
        addCatBtn.onclick = addNewCategory;
        appDiv.appendChild(addCatBtn);
        initSortable();
    }

    updateKPIs();
}

// =============================================
// VARIABLE INPUT MODAL
// =============================================
let _varInputResolve = null;

function showVarInputModal(varName, currentIdx, total) {
    return new Promise(resolve => {
        _varInputResolve = resolve;
        document.getElementById('var-input-name').textContent = '{{ ' + varName + ' }}';
        document.getElementById('varInputField').value = '';
        const counter = document.getElementById('var-input-counter');
        if (total > 1) {
            counter.textContent = `${currentIdx + 1} / ${total}`;
            counter.style.display = '';
        } else {
            counter.style.display = 'none';
        }
        const okBtn = document.getElementById('btn-var-input-ok');
        okBtn.textContent = currentIdx < total - 1 ? '次へ →' : '確定してコピー';
        document.getElementById('varInputModal').classList.remove('hidden');
        setTimeout(() => document.getElementById('varInputField').focus(), 60);
    });
}
function executeVarInput() {
    const val = document.getElementById('varInputField').value;
    const resolve = _varInputResolve;
    _varInputResolve = null;
    document.getElementById('varInputModal').classList.add('hidden');
    if (resolve) resolve(val);
}
function cancelVarInput() {
    const resolve = _varInputResolve;
    _varInputResolve = null;
    document.getElementById('varInputModal').classList.add('hidden');
    if (resolve) resolve(null);
}

// =============================================
// COPY WITH VARIABLES
// =============================================
async function copyTextWithVariables(cIndex, tIndex) {
    let text = appData[cIndex].templates[tIndex].content;
    const matches = text.match(/\{\{(.*?)\}\}/g);
    if (matches) {
        const uniqueVars = [...new Set(matches)];
        for (let i = 0; i < uniqueVars.length; i++) {
            const match = uniqueVars[i];
            const varName = match.replace(/[{}]/g, '');
            const val = await showVarInputModal(varName, i, uniqueVars.length);
            if (val === null) { showToast('コピーをキャンセルしました'); return; }
            const converted = convertToJpDatetime(val);
            text = text.split(match).join(converted || val);
        }
    }
    navigator.clipboard.writeText(text).then(() => {
        const tpl = appData[cIndex].templates[tIndex];
        tpl.copyCount = (tpl.copyCount || 0) + 1;
        addToHistory(tpl.title || '無題', text);
        showToast('コピーしました！');
        render();
        saveData(false, true);
    });
}

// =============================================
// HISTORY
// =============================================
function getHistory() { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); }
function addToHistory(title, text) {
    const history = getHistory();
    const now = new Date();
    const ts = `${now.getFullYear()}/${String(now.getMonth()+1).padStart(2,'0')}/${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
    history.unshift({ title, preview: text.slice(0,80), full: text, time: ts });
    if (history.length > 50) history.pop();
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    renderHistoryPanel();
}
function toggleHistory() { document.getElementById('historyPanel').classList.toggle('open'); renderHistoryPanel(); }
function renderHistoryPanel() {
    const list = document.getElementById('historyList');
    const h    = getHistory();
    if (!h.length) { list.innerHTML = '<div style="color:var(--dim);font-size:0.8rem;text-align:center;margin-top:40px;">まだコピー履歴がありません</div>'; return; }
    list.innerHTML = h.map((item, i) => `
        <div class="history-item" onclick="reCopyHistory(${i})">
            <div class="history-item-title">${esc(item.title)}</div>
            <div class="history-item-time">${item.time}</div>
            <div class="history-item-preview">${esc(item.preview)}${item.full.length > 80 ? '…' : ''}</div>
        </div>`).join('');
}
function reCopyHistory(i) {
    const h = getHistory()[i];
    if (!h) return;
    navigator.clipboard.writeText(h.full).then(() => showToast('履歴からコピーしました！'));
}
function clearHistory() {
    showConfirm('コピー履歴をすべて削除しますか？', () => {
        localStorage.removeItem(HISTORY_KEY); renderHistoryPanel(); showToast('履歴を削除しました');
    });
}

// =============================================
// COMMENT MEMO
// =============================================
function openCommentModal(c, t) {
    commentTarget = { c, t };
    document.getElementById('commentInput').value = appData[c].templates[t].comment || '';
    document.getElementById('commentModal').classList.remove('hidden');
}
function closeCommentModal() { document.getElementById('commentModal').classList.add('hidden'); }
function saveComment() {
    const { c, t } = commentTarget;
    appData[c].templates[t].comment = document.getElementById('commentInput').value;
    closeCommentModal(); render(); saveData(false, true); showToast('メモを保存しました');
}

// =============================================
// TEMPLATE CRUD
// =============================================
function saveTemplate(c, t) {
    const titleVal = document.getElementById(`edit-title-${c}-${t}`).value.trim();
    if (!titleVal) { showToast('⚠️ 件名を入力してください'); return; }
    const now = new Date();
    const ts  = `${now.getFullYear()}/${now.getMonth()+1}/${now.getDate()} ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;
    const tpl = appData[c].templates[t];
    tpl.title        = titleVal;
    tpl.tags         = document.getElementById(`edit-tags-${c}-${t}`).value.trim();
    tpl.content      = document.getElementById(`edit-content-${c}-${t}`).value;
    tpl.comment      = document.getElementById(`edit-comment-${c}-${t}`).value.trim();
    tpl.lastEditedBy = localStorage.getItem(USER_KEY) || "担当者";
    tpl.lastEditedAt = ts;
    unlockCurrentTemplate();
    editingState = { catIndex: null, tplIndex: null };
    render(); saveData(false, true);
}
function deleteTemplate(c, t) {
    showConfirm(`「${appData[c].templates[t].title || '無題'}」を削除しますか？`, () => {
        appData[c].templates.splice(t, 1); render(); saveData(false, true);
    });
}
async function startEdit(cIndex, tIndex) {
    const tpl  = appData[cIndex]?.templates[tIndex];
    const user = localStorage.getItem(USER_KEY) || '担当者';
    if (API_URL.startsWith("https://script.google.com") && tpl?.id) {
        try {
            const res = await Promise.race([
                fetch(API_URL, { method: 'POST', body: JSON.stringify({ action: 'lockTemplate', tplId: tpl.id, user }) }),
                new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 3000))
            ]);
            const data = await res.json();
            if (data.locked) {
                showConfirm(
                    `「${data.by}」さんが現在このテンプレートを編集中です。それでも編集を開始しますか？`,
                    () => { editingState = { catIndex: cIndex, tplIndex: tIndex }; render(); },
                    '強制編集', '⚠️ 編集中のテンプレート'
                );
                return;
            }
        } catch (_) {}
    }
    editingState = { catIndex: cIndex, tplIndex: tIndex };
    render();
}
function unlockCurrentTemplate() {
    const { catIndex, tplIndex } = editingState;
    if (catIndex === null || tplIndex === null) return;
    const tpl  = appData[catIndex]?.templates[tplIndex];
    const user = localStorage.getItem(USER_KEY) || '担当者';
    if (API_URL.startsWith("https://script.google.com") && tpl?.id) {
        fetch(API_URL, { method: 'POST', body: JSON.stringify({ action: 'unlockTemplate', tplId: tpl.id, user }) }).catch(() => {});
    }
}
function cancelEdit() { unlockCurrentTemplate(); editingState = { catIndex: null, tplIndex: null }; render(); }
function addNewTemplate(c) { appData[c].templates.push({ id:'t_'+Date.now(), title:'', tags:'', content:'', comment:'', copyCount:0 }); startEdit(c, appData[c].templates.length-1); saveData(false,true); }
function editCategory(c) {
    showInputModal('カテゴリ名変更', '新しいカテゴリ名', appData[c].title, n => {
        appData[c].title = n; render(); saveData(false, true);
    });
}
function deleteCategory(cIndex) {
    showConfirm(`カテゴリ「${appData[cIndex].title}」を中のテンプレートごと削除しますか？`, () => {
        appData.splice(cIndex, 1); render(); saveData(false, true, true);
    }, '⚠️ 削除する');
}
function addNewCategory() {
    showInputModal('カテゴリ追加', '新しいカテゴリー名', '', n => {
        if (!n.trim()) return;
        const newCat = { id: 'cat_' + Date.now(), title: n.trim(), templates: [] };
        appData.push(newCat);
        appData[appData.length - 1].templates.push({ id: 't_' + Date.now(), title: '', tags: '', content: '', comment: '', copyCount: 0 });
        render(); saveData(false, true); startEdit(appData.length - 1, 0);
    });
}
function toggleFavorite(id) {
    let favs = getFavorites();
    if (favs.includes(id)) { favs = favs.filter(f=>f!==id); showToast('お気に入りから解除しました'); }
    else { favs.push(id); showToast('お気に入りに追加しました'); }
    localStorage.setItem(FAV_KEY, JSON.stringify(favs)); render();
}

// =============================================
// AI
// =============================================
// =============================================
// AI SEARCH HELPERS
// =============================================
function extractSnippet(text, maxLen = 90) {
    if (!text) return '';
    // 変数プレースホルダーを除去、共通の挨拶行をスキップ
    const cleaned = text
        .replace(/\{\{[^}]+\}\}/g, '▲')
        .replace(/^いつも.*[\r\n]*/m, '')
        .replace(/^今後とも.*[\r\n]*/m, '')
        .trim();
    // 空でない行を取得し、先頭2行を結合
    const lines  = cleaned.split('\n').map(l => l.trim()).filter(l => l);
    if (!lines.length) return '';
    let result = lines[0];
    if (result.length < 45 && lines[1]) result += ' ' + lines[1];
    return result.length > maxLen ? result.slice(0, maxLen) + '…' : result;
}

async function aiSearch() {
    const query = document.getElementById('searchInput').value.trim();
    if (!query) { showToast('⚠️ 問い合わせ内容を入力してください'); return; }
    const btn = document.getElementById('btn-ai-search');
    const orig = btn.innerText; btn.innerText = '⏳ 思考中...'; btn.disabled = true;
    try {
        const allTpls = appData.flatMap(cat => cat.templates);
        const res     = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'findBestTemplate',
                text: query,
                templates: allTpls.map(t => ({ title: t.title, snippet: extractSnippet(t.content) }))
            })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        const results = data.results || [];
        if (results.length === 0) {
            const hint = data.debug ? `（AI応答: ${data.debug.slice(0, 60)}…）` : '';
            showToast('🤖 該当するテンプレートが見つかりませんでした' + hint, 5000);
            return;
        }
        openAiSearchModal(query, results);
    } catch(e) { showToast('❌ 通信エラーが発生しました', 3500); }
    finally { btn.innerText = orig; btn.disabled = false; }
}

function findTemplateByTitle(title) {
    for (let ci = 0; ci < appData.length; ci++) {
        const ti = appData[ci].templates.findIndex(t => t.title === title);
        if (ti !== -1) return { cIndex: ci, tIndex: ti };
    }
    return null;
}

function openAiSearchModal(query, results) {
    const preview    = query.length > 40 ? query.slice(0, 40) + '…' : query;
    document.getElementById('ai-search-query-preview').textContent = `「${preview}」への推薦結果`;

    const scoreColor = s => s >= 80 ? 'var(--success)' : s >= 60 ? 'var(--warn)' : 'var(--dim)';
    const rankClass  = i => ['rank-1','rank-2','rank-3','rank-other'][Math.min(i, 3)];

    document.getElementById('ai-search-result-list').innerHTML = results.map((r, i) => {
        const found   = findTemplateByTitle(r.title);
        const tpl     = found ? appData[found.cIndex].templates[found.tIndex] : null;
        const hasVars = tpl && /\{\{.*?\}\}/.test(tpl.content);
        const tagHtml = tpl && tpl.tags
            ? tpl.tags.split(',').map(t => `<span class="tag-badge" style="font-size:0.68rem;">🏷️ ${esc(t.trim())}</span>`).join('')
            : '';

        return `<div class="ai-search-card">
            <div class="ai-search-card-header">
                <div class="ai-search-rank ${rankClass(i)}">#${i + 1}</div>
                <div class="ai-search-body">
                    <div class="ai-search-title">${esc(r.title)}</div>
                    <div class="ai-search-score-row">
                        <div class="ai-search-bar-wrap">
                            <div class="ai-search-bar-fill" style="width:${r.score}%;background:${scoreColor(r.score)};"></div>
                        </div>
                        <span class="ai-search-score-num" style="color:${scoreColor(r.score)};">${r.score}点</span>
                    </div>
                    ${r.reason ? `<div class="ai-search-reason">${esc(r.reason)}</div>` : ''}
                </div>
            </div>
            ${tpl ? `
            <div class="ai-search-content">${esc(tpl.content)}</div>
            <div class="ai-search-card-footer">
                <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;">
                    <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
                        ${tagHtml}
                        ${hasVars ? `<span style="font-size:0.72rem;color:var(--dim);">⚙️ 変数あり</span>` : ''}
                        ${tpl.comment ? `<span style="font-size:0.72rem;color:var(--warn);">💬 ${esc(tpl.comment)}</span>` : ''}
                    </div>
                    <button class="btn btn-success btn-sm" onclick="copyFromAiSearch(${found.cIndex},${found.tIndex})">📋 コピー</button>
                </div>
            </div>` : `
            <div class="ai-search-card-footer">
                <span style="font-size:0.8rem;color:var(--danger);">⚠️ テンプレートが見つかりません（名前が変更された可能性があります）</span>
            </div>`}
        </div>`;
    }).join('');

    document.getElementById('aiSearchModal').classList.remove('hidden');
}

function closeAiSearchModal() {
    document.getElementById('aiSearchModal').classList.add('hidden');
}

async function copyFromAiSearch(cIndex, tIndex) {
    closeAiSearchModal();
    await copyTextWithVariables(cIndex, tIndex);
}
async function autoTag(cIndex, tIndex) {
    const content = document.getElementById(`edit-content-${cIndex}-${tIndex}`).value;
    if (!content.trim()) { showToast('⚠️ 先に本文を入力してください'); return; }
    const btn = document.getElementById(`btn-auto-tag-${cIndex}-${tIndex}`);
    const orig = btn.innerText; btn.innerText = "⏳"; btn.disabled = true;
    try {
        const res  = await fetch(API_URL, { method:"POST", body: JSON.stringify({ action:"generateTags", text:content }) });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        document.getElementById(`edit-tags-${cIndex}-${tIndex}`).value = data.result;
        showToast("✨ AIがタグを生成しました");
    } catch(e) { showToast('❌ 通信エラー: ' + e.message, 3500); }
    finally { btn.innerText = orig; btn.disabled = false; }
}

// =============================================
// CRANE
// =============================================
// =============================================
// GAS接続診断
// =============================================
async function runDiagnosis() {
    const resultDiv = document.getElementById('diagnosis-result');
    document.getElementById('diagnosisModal').classList.remove('hidden');
    resultDiv.innerHTML = '<span style="color:var(--dim)">🔄 テスト中...</span>';

    const row = (label, value, ok) => {
        const color = ok === true ? 'var(--success)' : ok === false ? 'var(--danger)' : 'var(--muted)';
        const icon  = ok === true ? '✅' : ok === false ? '❌' : 'ℹ️';
        return `<div style="display:flex;gap:10px;padding:6px 0;border-bottom:1px solid var(--border);">
            <span style="width:18px;flex-shrink:0;">${icon}</span>
            <span style="width:180px;flex-shrink:0;color:var(--dim);font-size:0.8rem;">${label}</span>
            <span style="color:${color};word-break:break-all;">${esc(String(value))}</span>
        </div>`;
    };

    const openBtn = `<button class="btn btn-ghost btn-sm" style="margin-top:10px;"
        onclick="window.open(API_URL,'_blank')">🔗 GAS URLをブラウザで直接開く</button>`;

    let html = openBtn + '<div style="margin-top:12px;">';

    // 1. URL形式チェック
    const urlOk = API_URL.startsWith('https://script.google.com/macros/s/') && API_URL.endsWith('/exec');
    html += row('API_URL形式', urlOk ? '正常' : API_URL.slice(0, 80), urlOk);

    if (!urlOk) {
        html += `</div><p style="color:var(--danger);margin-top:10px;">
            API_URLが正しい形式ではありません。<br>
            <code style="font-size:0.8em;">https://script.google.com/macros/s/＜ID＞/exec</code> の形式か確認してください。</p>`;
        resultDiv.innerHTML = html;
        return;
    }

    // 2. no-cors でサーバー到達確認（CORSエラーかネットワーク障害かを区別）
    let reachable = false;
    try {
        await fetch(API_URL, { mode: 'no-cors' });
        reachable = true;
    } catch (_) {
        reachable = false;
    }
    html += row('サーバー到達', reachable ? '可能' : '不可（ネットワーク障害の疑い）', reachable);

    if (!reachable) {
        html += `</div><p style="color:var(--danger);margin-top:10px;">
            GASサーバーに到達できません。<br>
            ・インターネット接続を確認してください<br>
            ・URLが正しいか確認してください（上の「GAS URLをブラウザで直接開く」で確認）</p>`;
        resultDiv.innerHTML = html;
        return;
    }

    // 3. 通常fetchでレスポンス取得（CORSチェック）
    let rawText = '', fetchOk = false, contentType = '—';
    try {
        const res   = await fetch(API_URL);
        fetchOk     = res.ok;
        contentType = res.headers.get('content-type') || '（なし）';
        rawText     = await res.text();
        html += row('CORS / HTTPステータス', res.status + ' ' + res.statusText, res.ok);
        html += row('Content-Type', contentType, contentType.includes('json'));
    } catch (e) {
        // サーバーには到達できるがCORSブロックされている
        html += row('CORS', 'ブロックされています: ' + e.message, false);
        html += `</div><p style="color:var(--danger);margin-top:10px;">
            <b>GASのデプロイ設定を確認してください：</b><br>
            1. GASエディタ → デプロイ → デプロイを管理<br>
            2. 対象デプロイの「編集（鉛筆アイコン）」をクリック<br>
            3. <b>「次のユーザーとして実行」→「自分（xxxx@gmail.com）」</b><br>
            4. <b>「アクセスできるユーザー」→「全員」</b>（Googleアカウント不要の方）<br>
            5. 「デプロイ」→承認画面が出たら許可する<br>
            6. 新しいURLをapp.jsの <code>API_URL</code> に貼り付ける</p>`;
        resultDiv.innerHTML = html;
        return;
    }

    // 4. JSONパースチェック
    let parsed = null;
    try {
        parsed = JSON.parse(rawText);
        html += row('JSONパース', '成功', true);
    } catch (_) {
        html += row('JSONパース', '失敗', false);
        html += row('レスポンス先頭', rawText.slice(0, 150), null);
        html += `</div><p style="color:var(--danger);margin-top:10px;">
            GASがHTMLを返しています（ログインページにリダイレクトされている可能性）。<br>
            上の「GAS URLをブラウザで直接開く」でJSONが表示されるか確認してください。<br>
            HTMLが表示される場合 → デプロイの承認が完了していません。</p>`;
        resultDiv.innerHTML = html;
        return;
    }

    // 5. GASエラーフィールドチェック
    if (parsed.error) {
        html += row('GASスクリプトエラー', parsed.error, false);
        html += `</div><p style="color:var(--danger);margin-top:10px;">
            GASスクリプト内でエラーが発生しています。<br>
            よくある原因：スプレッドシートに「Data」シートが存在しない。<br>
            GASエディタの「実行ログ」で詳細を確認してください。</p>`;
        resultDiv.innerHTML = html;
        return;
    }

    // 6. データ構造チェック
    const hasJson = parsed.json !== undefined;
    html += row('データフィールド', hasJson ? '正常' : '異常（jsonフィールドなし）', hasJson);
    if (hasJson) {
        try {
            const data = JSON.parse(parsed.json);
            html += row('カテゴリ数', data.length + ' 件', true);
        } catch(_) {
            html += row('データパース', '失敗', false);
        }
    }
    html += row('最終更新タイムスタンプ', parsed.updatedAt || '0（未保存）', null);

    html += '</div>';
    if (hasJson) {
        html += `<p style="color:var(--success);margin-top:12px;font-weight:700;">
            ✅ GASとの接続は正常です。ページをリロードして再同期してください。</p>`;
    }

    resultDiv.innerHTML = html;
}

function toggleCrane(element) {
    if (element.classList.contains('active')) return;
    const items = ['📄','📄','📄','📄','📄','🔥','💎','✨','🎁','👑'];
    document.getElementById('treasure-icon').innerText = items[Math.floor(Math.random()*items.length)];
    element.classList.add('active');
    setTimeout(() => element.classList.remove('active'), 3500);
}

// =============================================
// SORTABLE
// =============================================
function initSortable() {
    const tocList = document.getElementById('toc-list');
    if (tocList) {
        new Sortable(tocList, { animation:150, ghostClass:'sortable-ghost', onEnd:function(evt) {
            if (evt.oldIndex===evt.newIndex) return;
            const item = appData.splice(evt.oldIndex,1)[0];
            appData.splice(evt.newIndex,0,item);
            render(); saveData(false,true);
        }});
    }
    document.querySelectorAll('tbody').forEach(tbody => {
        new Sortable(tbody, { animation:150, handle:'th', draggable:'tr', ghostClass:'sortable-ghost', onEnd:function(evt) {
            if (evt.oldIndex===evt.newIndex) return;
            const cIdx = parseInt(evt.from.dataset.catIndex);
            const item = appData[cIdx].templates.splice(evt.oldIndex,1)[0];
            appData[cIdx].templates.splice(evt.newIndex,0,item);
            render(); saveData(false,true);
        }});
    });
}

// =============================================
// MISC
// =============================================
// =============================================
// AI REPLY GENERATION
// =============================================
const AI_STEPS = [
    { icon: '💬', label: '問い合わせを解析', desc: '内容・意図・緊急度を確認' },
    { icon: '📖', label: '仕様書を参照',     desc: 'バグ状況・ゲーム仕様を照合' },
    { icon: '✍️', label: '返信文を生成',     desc: '文体・規定に合わせて作成' }
];

function openAiReplyModal() {
    resetAiReply();
    document.getElementById('aiReplyModal').classList.remove('hidden');
    setTimeout(() => document.getElementById('ai-inquiry-input').focus(), 60);
}
function closeAiReplyModal() {
    document.getElementById('aiReplyModal').classList.add('hidden');
}
function resetAiReply() {
    document.getElementById('ai-panel-input').style.display    = '';
    document.getElementById('ai-panel-progress').style.display = 'none';
    document.getElementById('ai-panel-result').style.display   = 'none';
}

function renderAiSteps(activeIndex) {
    document.getElementById('ai-progress-steps').innerHTML = AI_STEPS.map((s, i) => {
        const state = i < activeIndex ? 'done' : i === activeIndex ? 'active' : 'pending';
        const icon  = state === 'done' ? '✅' : state === 'active' ? s.icon : '◯';
        const stat  = state === 'done' ? '完了' : state === 'active' ? '処理中' : '';
        return `<div class="ai-step ${state}">
            <span class="ai-step-icon">${icon}</span>
            <div style="flex:1;">
                <div class="ai-step-label">${s.label}</div>
                <div class="ai-step-desc">${s.desc}</div>
            </div>
            <span class="ai-step-status">${stat}</span>
        </div>`;
    }).join('');
}

async function runAiReply() {
    const inquiry = document.getElementById('ai-inquiry-input').value.trim();
    if (!inquiry) { showToast('問い合わせ文を入力してください'); return; }

    document.getElementById('ai-panel-input').style.display    = 'none';
    document.getElementById('ai-panel-progress').style.display = '';
    document.getElementById('ai-panel-result').style.display   = 'none';

    let animStep = 0;
    renderAiSteps(animStep);

    // ステップアニメーション（700ms刻み）
    const stepTimer = setInterval(() => {
        if (animStep < AI_STEPS.length - 1) { animStep++; renderAiSteps(animStep); }
    }, 700);

    try {
        // API呼び出しと最低表示時間を並行実行
        const [result] = await Promise.all([
            fetch(API_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'generateReply', inquiry })
            }).then(r => r.json()),
            new Promise(r => setTimeout(r, AI_STEPS.length * 700 + 200))
        ]);

        clearInterval(stepTimer);
        renderAiSteps(AI_STEPS.length);

        await new Promise(r => setTimeout(r, 350));
        showAiResult(result);

    } catch (e) {
        clearInterval(stepTimer);
        showAiError('通信エラーが発生しました。サーバー接続を確認してください。');
    }
}

function showAiResult(result) {
    document.getElementById('ai-panel-progress').style.display = 'none';
    const panel = document.getElementById('ai-panel-result');

    if (result.error) { showAiError(result.error); return; }

    const resultText = result.reply || '';
    const source     = result.specUsed ? `仕様書「${esc(result.specUsed)}」を参照して生成` : '仕様書をもとにAIが生成';

    panel.innerHTML = `
        <div class="ai-badge gen">✨ AI生成返信</div>
        <div class="ai-source-note">${source}しました。内容を確認してからご利用ください。</div>
        <div class="ai-result-box">${esc(resultText)}</div>
        <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:6px;">
            <button class="btn btn-success" onclick="copyAiResult()">📋 コピー</button>
            <button class="btn btn-purple btn-sm" onclick="saveAiReplyAsTemplate()">💾 テンプレとして保存</button>
            <button class="btn btn-ghost" onclick="resetAiReply()">↩ 再入力</button>
        </div>`;
    panel.dataset.resultText  = resultText;
    panel.dataset.resultTitle = '';
    panel.style.display = '';
}

function showAiError(msg) {
    document.getElementById('ai-panel-progress').style.display = 'none';
    const panel = document.getElementById('ai-panel-result');
    panel.innerHTML = `
        <div class="ai-badge err">❌ エラー</div>
        <div style="font-size:0.88rem; color:var(--muted); margin:8px 0 14px; line-height:1.6;">${esc(msg)}</div>
        <button class="btn btn-ghost" onclick="resetAiReply()">↩ 再入力</button>`;
    panel.style.display = '';
}

function copyAiResult() {
    const text = document.getElementById('ai-panel-result').dataset.resultText || '';
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
        addToHistory('AI生成返信', text);
        showToast('コピーしました！');
    });
}

let _saveTplText = '';

function saveAiReplyAsTemplate() {
    const text = document.getElementById('ai-panel-result').dataset.resultText || '';
    if (!text) return;
    _saveTplText = text;
    document.getElementById('saveTplTitle').value = '';
    renderSaveTplCatList();
    closeAiReplyModal();
    document.getElementById('saveTplModal').classList.remove('hidden');
    setTimeout(() => document.getElementById('saveTplTitle').focus(), 60);
}
function closeSaveTplModal() {
    document.getElementById('saveTplModal').classList.add('hidden');
}
function renderSaveTplCatList() {
    document.getElementById('saveTplCatList').innerHTML = appData.map((cat, i) => `
        <div class="cat-select-item" onclick="executeSaveTpl(${i})">
            <span class="cat-select-name">${esc(cat.title)}</span>
            <span class="cat-select-count">${cat.templates.length}件</span>
        </div>`).join('');
}
function executeSaveTpl(catIndex) {
    const title = document.getElementById('saveTplTitle').value.trim();
    if (!title) { showToast('⚠️ 件名を入力してください'); document.getElementById('saveTplTitle').focus(); return; }
    appData[catIndex].templates.push({
        id: 't_' + Date.now(), title,
        tags: 'AI生成', content: _saveTplText, comment: '', copyCount: 0
    });
    render(); saveData(false, true);
    closeSaveTplModal();
    showToast(`「${title}」を保存しました`);
}

// =============================================
// SPEC MANAGEMENT
// =============================================
async function loadSpec() {
    if (!API_URL.startsWith("https://script.google.com")) return;
    try {
        const res  = await fetch(API_URL + '?action=getSpec');
        const data = await res.json();
        if (Array.isArray(data.spec)) specData = data.spec;
    } catch(e) {
        console.log('[BH-CS] 仕様書読み込みエラー:', e.message);
    }
}

function openSpecModal() {
    renderSpecTable();
    document.getElementById('specModal').classList.remove('hidden');
}
function closeSpecModal() {
    document.getElementById('specModal').classList.add('hidden');
}

function renderSpecTable() {
    const wrap = document.getElementById('spec-table-wrap');
    if (!wrap) return;
    if (specData.length === 0) {
        wrap.innerHTML = `<div style="text-align:center; color:var(--dim); padding:28px 0; font-size:0.88rem;">
            仕様書にデータがありません。「行追加」ボタンから追加してください。</div>`;
        return;
    }
    let html = `<table class="spec-table">
        <thead><tr>
            <th style="width:22%;">項目名</th>
            <th style="width:46%;">内容・状況</th>
            <th style="width:22%;">備考</th>
            <th style="width:10%;"></th>
        </tr></thead><tbody>`;
    specData.forEach((row, i) => {
        html += `<tr>
            <td><input class="spec-input" value="${esc(row.item||'')}"
                oninput="specData[${i}].item=this.value"
                placeholder="例: ログイン不具合 E-4023"></td>
            <td><textarea class="spec-input" style="height:72px;"
                oninput="specData[${i}].content=this.value"
                placeholder="例: 現在調査中">${esc(row.content||'')}</textarea></td>
            <td><input class="spec-input" value="${esc(row.note||'')}"
                oninput="specData[${i}].note=this.value"
                placeholder="任意"></td>
            <td style="text-align:center; vertical-align:middle;">
                <button class="btn btn-danger btn-xs" onclick="deleteSpecRow(${i})">🗑️</button>
            </td>
        </tr>`;
    });
    html += `</tbody></table>`;
    wrap.innerHTML = html;
}

function addSpecRow() {
    specData.push({ id: 'spec_' + Date.now(), item: '', content: '', note: '' });
    renderSpecTable();
    // フォーカスを最後の行の「項目名」欄へ
    const inputs = document.querySelectorAll('#spec-table-wrap .spec-input');
    if (inputs.length >= 3) inputs[inputs.length - 3].focus();
}

function deleteSpecRow(i) {
    specData.splice(i, 1);
    renderSpecTable();
    showToast('行を削除しました（Driveに保存するまで確定しません）');
}

async function saveSpec() {
    if (!API_URL.startsWith("https://script.google.com")) {
        showToast('❌ オフラインのため保存できません');
        return;
    }
    const btn  = document.getElementById('btn-save-spec');
    const orig = btn.innerText;
    btn.innerText = '⏳ 保存中...';
    btn.disabled  = true;
    try {
        const res  = await fetch(API_URL, {
            method: 'POST',
            body:   JSON.stringify({ action: 'saveSpec', spec: specData })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        showToast('✅ 仕様書をDriveに保存しました');
    } catch(e) {
        showToast('❌ 保存エラー: ' + e.message);
    } finally {
        btn.innerText = orig;
        btn.disabled  = false;
    }
}

// =============================================
// KNOWLEDGE COLLECTION
// =============================================
function openKnowledgeModal() {
    knowledgeProposed = [];
    document.getElementById('kn-panel-settings').style.display = '';
    document.getElementById('kn-panel-loading').style.display  = 'none';
    document.getElementById('kn-panel-review').style.display   = 'none';
    document.getElementById('kn-modal-footer').style.display   = 'none';
    document.getElementById('knowledgeModal').classList.remove('hidden');
}
function closeKnowledgeModal() {
    document.getElementById('knowledgeModal').classList.add('hidden');
}

async function runKnowledgeBuild() {
    const days = document.getElementById('kn-days').value;
    document.getElementById('kn-panel-settings').style.display = 'none';
    document.getElementById('kn-panel-loading').style.display  = '';
    document.getElementById('kn-modal-footer').style.display   = 'none';

    const msgs = [
        'HelpShiftからチケットを取得中...',
        'AIがやり取りを解析中...',
        '仕様書エントリを抽出中...'
    ];
    let mi = 0;
    const msgTimer = setInterval(() => {
        document.getElementById('kn-loading-msg').textContent = msgs[Math.min(++mi, msgs.length - 1)];
    }, 2500);

    try {
        const res  = await fetch(API_URL, {
            method: 'POST',
            body:   JSON.stringify({ action: 'buildKnowledge', days: parseInt(days) })
        });
        const data = await res.json();
        clearInterval(msgTimer);

        if (data.error) throw new Error(data.error);

        knowledgeProposed = data.items || [];
        document.getElementById('kn-panel-loading').style.display = 'none';
        document.getElementById('kn-panel-review').style.display  = '';

        const scanned  = data.scanned  || 0;
        const found    = knowledgeProposed.length;
        const header   = document.getElementById('kn-review-header');
        header.innerHTML = `<b>${scanned}</b>件のチケットを解析 → <b style="color:${found > 0 ? 'var(--success)' : 'var(--dim)'};">${found}</b>件の新規ナレッジを抽出しました。`;

        if (found === 0) {
            const debugInfo = (data.debug_errors && data.debug_errors.length)
                ? `<details style="margin-top:12px; text-align:left;"><summary style="font-size:0.75rem; color:var(--dim); cursor:pointer;">🔍 デバッグ情報</summary><pre style="font-size:0.72rem; color:var(--dim); margin-top:6px; white-space:pre-wrap;">${esc(data.debug_errors.join('\n'))}</pre></details>`
                : '';
            document.getElementById('kn-review-list').innerHTML =
                `<div style="text-align:center; color:var(--dim); padding:24px; font-size:0.88rem;">新しいナレッジは見つかりませんでした。<br>期間を広げるか、チケットが解決済みになっているか確認してください。${debugInfo}</div>`;
        } else {
            renderKnowledgeReview();
            document.getElementById('kn-modal-footer').style.display = 'flex';
        }

    } catch(e) {
        clearInterval(msgTimer);
        document.getElementById('kn-panel-loading').style.display = 'none';
        document.getElementById('kn-panel-review').style.display  = '';
        document.getElementById('kn-review-list').innerHTML =
            `<div style="color:var(--danger); font-size:0.88rem; padding:8px 0;">${esc(e.message)}</div>
             <div style="margin-top:12px;"><button class="btn btn-ghost btn-sm" onclick="document.getElementById('kn-panel-settings').style.display='';document.getElementById('kn-panel-review').style.display='none';">← 設定に戻る</button></div>`;
    }
}

function renderKnowledgeReview() {
    document.getElementById('kn-review-list').innerHTML = knowledgeProposed.map((item, i) => `
        <div class="kn-item" onclick="this.querySelector('.kn-checkbox').click()">
            <label style="display:flex; gap:12px; align-items:flex-start; cursor:pointer;" onclick="event.stopPropagation()">
                <input type="checkbox" class="kn-checkbox" data-index="${i}" checked
                    style="margin-top:3px; flex-shrink:0; width:15px; height:15px; accent-color:var(--primary);">
                <div style="flex:1;">
                    <div class="kn-item-title">${esc(item.item)}<span class="kn-badge">自動収集</span></div>
                    <div class="kn-item-content">${esc(item.content)}</div>
                    ${item.note ? `<div class="kn-item-note">${esc(item.note)}</div>` : ''}
                </div>
            </label>
        </div>`).join('');
}

function knSelectAll(checked) {
    document.querySelectorAll('.kn-checkbox').forEach(cb => cb.checked = checked);
}

async function addSelectedKnowledge() {
    const selected = [];
    document.querySelectorAll('.kn-checkbox').forEach(cb => {
        if (cb.checked) selected.push(knowledgeProposed[parseInt(cb.dataset.index)]);
    });
    if (selected.length === 0) { showToast('⚠️ 項目を選択してください'); return; }

    const btn  = document.getElementById('btn-add-kn');
    const orig = btn.innerText;
    btn.disabled  = true;
    btn.innerText = '⏳ 追加中...';

    try {
        const res  = await fetch(API_URL, {
            method: 'POST',
            body:   JSON.stringify({ action: 'addKnowledgeItems', items: selected })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);

        specData = specData.concat(selected);
        showToast(`✅ ${selected.length}件のナレッジを仕様書に追加しました`);
        closeKnowledgeModal();
        if (!document.getElementById('specModal').classList.contains('hidden')) renderSpecTable();

    } catch(e) {
        showToast('❌ 追加エラー: ' + e.message);
    } finally {
        btn.disabled  = false;
        btn.innerText = orig;
    }
}

// =============================================
// CONFLICT MODAL
// =============================================
function openConflictModal() {
    const draft = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null');
    if (!draft) return;
    document.getElementById('conflict-draft-time').textContent  = draft.savedAt || '—';
    const diff = computeDiff(draft.data || [], appData);
    renderConflictDiff(diff);
    document.getElementById('conflictModal').classList.remove('hidden');
}
function closeConflictModal() {
    document.getElementById('conflictModal').classList.add('hidden');
}
function computeDiff(draftCats, serverCats) {
    const result = [];
    const serverMap = {};
    serverCats.forEach(cat => {
        (cat.templates || []).forEach(tpl => { if (tpl.id) serverMap[tpl.id] = tpl; });
    });
    draftCats.forEach(cat => {
        (cat.templates || []).forEach(dTpl => {
            if (!dTpl.id) return;
            const sTpl = serverMap[dTpl.id];
            if (!sTpl) {
                result.push({ type: 'added', title: dTpl.title, draftVal: dTpl.content, serverVal: null });
            } else if (dTpl.content !== sTpl.content || dTpl.title !== sTpl.title) {
                result.push({ type: 'modified', title: dTpl.title, draftVal: dTpl.content, serverVal: sTpl.content });
            }
        });
    });
    return result;
}
function renderConflictDiff(diff) {
    const container = document.getElementById('conflict-diff-list');
    if (!diff.length) {
        container.innerHTML = '<div style="color:var(--dim);font-size:0.85rem;text-align:center;padding:16px;">差分はありません（データ構造の競合）</div>';
        return;
    }
    container.innerHTML = diff.map(d => `
        <div class="conflict-diff-item">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
                <span class="conflict-diff-badge ${d.type === 'added' ? 'badge-added' : 'badge-modified'}">${d.type === 'added' ? '追加' : '変更'}</span>
                <span style="font-weight:700; font-size:0.88rem;">${esc(d.title)}</span>
            </div>
            <div class="conflict-version-grid">
                <div class="conflict-version-card draft">
                    <div class="conflict-version-label">✏️ 自分の編集</div>
                    <div class="conflict-version-body">${esc((d.draftVal || '').slice(0, 200))}${(d.draftVal || '').length > 200 ? '…' : ''}</div>
                </div>
                ${d.serverVal !== null ? `
                <div class="conflict-version-card server">
                    <div class="conflict-version-label">☁️ サーバー版</div>
                    <div class="conflict-version-body">${esc(d.serverVal.slice(0, 200))}${d.serverVal.length > 200 ? '…' : ''}</div>
                </div>` : ''}
            </div>
        </div>`).join('');
}
function conflictForceSave() {
    const draft = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null');
    if (!draft) { closeConflictModal(); return; }
    appData = draft.data;
    localStorage.removeItem(DRAFT_KEY);
    closeConflictModal();
    executeSave(true, false);
    showToast('自分の編集内容で上書き保存しました');
}
function conflictUseServer() {
    localStorage.removeItem(DRAFT_KEY);
    closeConflictModal();
    init();
    showToast('サーバー版を読み込みました');
}

// =============================================
// CONFIRM MODAL
// =============================================
function showConfirm(message, onOk, okLabel = '削除', title = '確認') {
    document.getElementById('confirmMessage').textContent = message;
    document.getElementById('confirmTitle').textContent   = title;
    document.getElementById('confirmOkBtn').textContent   = okLabel;
    confirmCallback = onOk;
    document.getElementById('confirmModal').classList.remove('hidden');
}
function closeConfirmModal() {
    document.getElementById('confirmModal').classList.add('hidden');
    confirmCallback = null;
}
function executeConfirm() { const cb = confirmCallback; closeConfirmModal(); if (cb) cb(); }

// =============================================
// INPUT MODAL
// =============================================
function showInputModal(title, label, defaultVal, onOk) {
    document.getElementById('inputModalTitle').textContent = title;
    document.getElementById('inputModalLabel').textContent = label;
    document.getElementById('inputModalField').value       = defaultVal || '';
    inputCallback = onOk;
    document.getElementById('inputModal').classList.remove('hidden');
    setTimeout(() => document.getElementById('inputModalField').select(), 50);
}
function closeInputModal() {
    document.getElementById('inputModal').classList.add('hidden');
    inputCallback = null;
}
function executeInput() {
    const val = document.getElementById('inputModalField').value.trim();
    const cb  = inputCallback;
    closeInputModal();
    if (cb && val) cb(val);
}

function showToast(m, duration=2200) { const t=document.getElementById('toast'); t.innerText=m; t.style.opacity=1; setTimeout(()=>t.style.opacity=0,duration); }
function forceSync() {
    if (editingState.catIndex !== null) {
        showConfirm('編集中の内容が破棄されます。続行しますか？', () => {
            editingState = { catIndex: null, tplIndex: null };
            init(); showToast('最新状態に更新しました');
        }, '続行', '⚠️ 確認');
        return;
    }
    init(); showToast('最新状態に更新しました');
}
function exportData() {
    const d  = new Date();
    const ds = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
    const a  = Object.assign(document.createElement('a'), {
        href: URL.createObjectURL(new Blob([JSON.stringify(appData, null, 2)], { type: 'application/json' })),
        download: `backup_${ds}.json`
    });
    a.click();
}
function handleFileSelect(e) {
    const r = new FileReader();
    r.onload = ev => {
        try {
            const parsed = JSON.parse(ev.target.result);
            if (!Array.isArray(parsed)) throw new Error("配列ではありません");
            const isValid = parsed.every(cat =>
                cat && typeof cat === 'object' &&
                typeof cat.title === 'string' &&
                Array.isArray(cat.templates)
            );
            if (!isValid) throw new Error("データ構造が不正です");
            appData = parsed;
            ensureIds();
            saveData(true);
            render();
            showToast('読み込みが完了しました');
        } catch(err) {
            showToast('❌ 読み込みエラー: ' + err.message, 4000);
        }
    };
    r.readAsText(e.target.files[0]);
    e.target.value = '';
}
window.onscroll = () => { document.getElementById("btn-back-to-top").style.display = window.scrollY>300?"flex":"none"; };
function scrollToTop() { window.scrollTo({ top:0, behavior:'smooth' }); }

// =============================================
// THEME TOGGLE
// =============================================
const THEME_KEY = 'bh_cs_theme';
function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    const isLight = theme === 'light';
    const track   = document.getElementById('ttTrack');
    const icon    = document.getElementById('themeIcon');
    const label   = document.getElementById('themeLabel');
    if (track)  track.classList.toggle('on', isLight);
    if (icon)   icon.textContent  = isLight ? '☀️' : '🌙';
    if (label)  label.textContent = isLight ? 'ライト' : 'ダーク';
    localStorage.setItem(THEME_KEY, theme);
}
function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    applyTheme(current === 'dark' ? 'light' : 'dark');
}
// テーマを復元
(function() {
    const saved = localStorage.getItem('bh_cs_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', saved);
})();

window.onload = function() {
    const saved = localStorage.getItem(THEME_KEY) || 'dark';
    applyTheme(saved);
    init();
    loadSpec();

    // Ctrl+/ で検索ボックスにフォーカス
    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === '/') {
            e.preventDefault();
            const si = document.getElementById('searchInput');
            si.focus(); si.select();
        }
        if (e.key === 'Escape') {
            cancelVarInput();
            closeAiSearchModal();
            closeConflictModal();
            closeConfirmModal();
            closeInputModal();
            closeCommentModal();
            closeSaveTplModal();
            closeAiReplyModal();
            closeSpecModal();
            closeKnowledgeModal();
            document.getElementById('diagnosisModal').classList.add('hidden');
        }
    });

    // 入力モーダルのEnterキー確定
    document.getElementById('inputModalField').addEventListener('keydown', e => {
        if (e.key === 'Enter') executeInput();
    });

    // 変数入力モーダルのEnterキー確定
    document.getElementById('varInputField').addEventListener('keydown', e => {
        if (e.key === 'Enter') executeVarInput();
    });
};

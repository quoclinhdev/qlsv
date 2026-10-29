/* Trợ lý AI-native: thanh lệnh (Ctrl+K), hiểu câu tiếng Việt tự nhiên, bộ lọc dạng chip, gợi ý phân tích.
   Chạy hoàn toàn trên trình duyệt (luật + dữ liệu đang hiển thị); không gửi dữ liệu ra ngoài. */
(function () {
  const qs = s => document.querySelector(s);
  const dlg = qs('#aiDlg'); if (!dlg) return;
  const inp = qs('#aiInput'), prev = qs('#aiPrev'), sugg = qs('#aiSugg'), bar = qs('#aiActive'),
        panel = qs('#aiPanel'), list = qs('#aiList'), trigger = qs('#aiOpen');
  const esc = s => String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
  const STOP = new Set(('xoa diem cho toi minh hay giup tim loc hien thi xem nhung cac sinh vien sv nay nao co la cua ra ' +
    'tat ca het duoc va ma o trong de voi nhu the tinh trang hoc ky nam danh sach ai hoac tu').split(' '));
  let cfg = null, applied = {filters: [], query: ''}, token = 0;

  /* ---------- hiểu câu lệnh ---------- */
  function parse(raw) {
    let s = norm(raw).toLowerCase().replace(/[“”"'.,;!?]+$/g, '').replace(/\s+/g, ' ').trim();
    const out = {filters: [], cmd: null};
    const eat = (re, fn) => {
      const m = s.match(re); if (!m) return false;
      s = (s.slice(0, m.index) + ' ' + s.slice(m.index + m[0].length)).replace(/\s+/g, ' ').trim(); fn(m); return true;
    };
    const A = cfg.actions || {};
    if (A.delete && eat(/\b(xoa|go bo|loai bo)\b/, () => out.cmd = 'delete')) {}
    else if (eat(/\b(bo chon|huy chon)\b/, () => out.cmd = 'clear')) {}
    else if (eat(/\bchon\b/, () => out.cmd = 'select')) {}
    else if (A.export && eat(/\bxuat\b(?! sac)( ra| file| excel| xlsx)*/, () => out.cmd = 'export')) {}
    else if (A.add && eat(/\bthem\b( moi)?/, () => out.cmd = 'add')) {}
    else if (eat(/\b(tom tat|phan tich|tong quan|nhan xet|goi y)\b/, () => out.cmd = 'insight')) {}

    if (cfg.numLabel) {  // so sánh số: dưới / trên / từ..đến / trở lên / bằng
      const N = '(\\d+(?:[.,]\\d+)?)', L = cfg.numLabel, f = x => parseFloat(x.replace(',', '.'));
      const add = (lab, test) => out.filters.push({label: `${L} ${lab}`, test: i => i.num != null && test(i.num)});
      eat(new RegExp(`tu\\s*${N}\\s*(?:den|toi|-)\\s*${N}`), m => { const a = f(m[1]), b = f(m[2]); add(`${a}–${b}`, v => v >= a && v <= b); });
      eat(new RegExp(`${N}\\s*tro len`), m => { const a = f(m[1]); add(`≥ ${a}`, v => v >= a); });
      eat(new RegExp(`${N}\\s*tro xuong`), m => { const a = f(m[1]); add(`≤ ${a}`, v => v <= a); });
      eat(new RegExp(`(?:duoi|nho hon|thap hon|<)\\s*${N}`), m => { const a = f(m[1]); add(`< ${a}`, v => v < a); });
      eat(new RegExp(`(?:tren|lon hon|cao hon|>)\\s*${N}`), m => { const a = f(m[1]); add(`> ${a}`, v => v > a); });
      eat(new RegExp(`(?:bang|=)\\s*${N}`), m => { const a = f(m[1]); add(`= ${a}`, v => v === a); });
    }
    (cfg.rules || []).forEach(r => eat(r.re, () => out.filters.push({label: r.label, group: r.group, test: r.test})));
    if (cfg.has && cfg.has.lop) eat(/\blop\s+([a-z0-9_.\-]+)/, m =>
      out.filters.push({label: 'Lớp ' + m[1].toUpperCase(), test: i => norm(i.lop || '').toLowerCase().includes(m[1])}));
    if (cfg.has && cfg.has.khoaHoc) eat(/\b(?:khoa|k)\s*(\d{1,3})\b/, m =>
      out.filters.push({label: 'Khóa ' + m[1], test: i => (/\d+/.exec(i.khoaHoc || '') || [])[0] === m[1]}));

    const all = s.split(' ').filter(Boolean), extra = cfg.stop || [];
    let tokens = all.filter(w => !STOP.has(w) && !extra.includes(w));
    if (!tokens.length && !out.cmd && !out.filters.length) tokens = all;
    if (tokens.length) out.filters.push({label: `Chứa “${tokens.join(' ')}”`, test: i => tokens.every(w => i.text.includes(w))});
    return out;
  }
  function match(filters, i) {  // cùng nhóm (group) thì HOẶC, khác nhóm thì VÀ
    const g = {};
    for (const f of filters) {
      if (f.group) (g[f.group] = g[f.group] || []).push(f);
      else if (!f.test(i)) return false;
    }
    return Object.values(g).every(a => a.some(f => f.test(i)));
  }
  const infos = () => cfg.table.rows().nodes().toArray().map(tr => ({tr, i: cfg.info(tr)}));

  $.fn.dataTable.ext.search.push(function (settings, data, idx) {
    if (!cfg || !applied.filters.length || settings.nTable !== cfg.table.table().node()) return true;
    const tr = settings.aoData[idx].nTr; return match(applied.filters, cfg.info(tr));
  });

  /* ---------- hiển thị ---------- */
  function renderActive() {
    if (!cfg || !applied.filters.length) { bar.hidden = true; return; }
    const shown = cfg.table.rows({search: 'applied'}).count(), tot = cfg.table.rows().count();
    bar.hidden = false;
    bar.innerHTML = '<span class="spark">✦</span><span class="ai-q">“' + esc(applied.query) + '”</span>' +
      applied.filters.map((f, k) => '<span class="chip">' + esc(f.label) + ' <button data-i="' + k + '" title="Bỏ bộ lọc này">✕</button></span>').join('') +
      '<span class="ai-count"><b>' + shown + '</b> / ' + tot + ' ' + esc(cfg.noun) + '</span><button class="btn gray sm" data-clear>Bỏ lọc AI</button>';
  }
  bar.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.hasAttribute('data-clear')) applied.filters = [];
    else applied.filters.splice(+b.dataset.i, 1);
    cfg.table.draw(); renderActive();
  });

  function apply(filters, query) { applied = {filters, query}; cfg.table.draw(); renderActive(); }
  const filteredRows = () => cfg.table.rows({search: 'applied'}).nodes().toArray();
  const setChecked = (nodes, v) => { $(nodes).find('input.chk').prop('checked', v); cfg.bulk && cfg.bulk.update(); };

  function preview(raw) {
    if (!raw.trim()) {
      prev.innerHTML = '<div class="ai-note">Gõ yêu cầu bằng tiếng Việt, ví dụ <i>“' + esc((cfg.examples || ['điểm dưới 50'])[0]) +
        '”</i>. Trợ lý sẽ hiểu và lọc bảng cho bạn.</div>'; return;
    }
    const p = parse(raw), n = infos().filter(x => match(p.filters, x.i)).length;
    const act = {delete: `Sẽ <b>xóa</b> ${n} ${esc(cfg.noun)} (cần bạn xác nhận)`, select: `Sẽ <b>chọn</b> ${n} ${esc(cfg.noun)}`,
      clear: 'Sẽ bỏ chọn tất cả', export: 'Sẽ xuất Excel', add: 'Sẽ mở form thêm mới', insight: 'Sẽ mở phân tích tổng quan'}[p.cmd];
    prev.innerHTML = (p.filters.length ? '<div>Đã hiểu: ' + p.filters.map(f => '<span class="chip">' + esc(f.label) + '</span>').join('') + '</div>' : '') +
      '<div class="ai-note">' + (act ? act + ' · ' : '') + '<b>' + n + '</b> ' + esc(cfg.noun) + ' phù hợp. Nhấn <kbd>Enter</kbd> để thực hiện.</div>';
  }

  function askDelete(rows, noFilter) {
    const n = rows.length;
    if (!n) { prev.innerHTML = '<div class="ai-note">Không có dòng nào phù hợp để xóa.</div>'; return; }
    const names = rows.slice(0, 5).map(x => esc(x.tr.dataset.mssv + ' – ' + (x.tr.cells[3].textContent + ' ' + x.tr.cells[4].textContent).trim()));
    prev.innerHTML = '<div class="ai-confirm"><div class="warn">' + esc(cfg.deleteVerb) + ' ' + n + ' ' + esc(cfg.noun) + '?' +
      (noFilter ? ' (KHÔNG có bộ lọc – sẽ xóa toàn bộ danh sách đang hiển thị)' : '') + '</div><ul>' +
      names.map(t => '<li>' + t + '</li>').join('') + (n > 5 ? '<li>… và ' + (n - 5) + ' ' + esc(cfg.noun) + ' khác</li>' : '') +
      '</ul><button class="btn red" id="aiYes">Xác nhận xóa</button> <button class="btn gray" id="aiNo">Hủy</button></div>';
    qs('#aiNo').onclick = () => dlg.close();
    qs('#aiYes').onclick = async () => {
      const ds = rows.map(x => x.tr.dataset.mssv); dlg.close();
      await cfg.actions.delete(ds); renderActive(); renderInsights(false);
    };
  }

  function execute(raw) {
    if (!cfg) return;
    if (!raw.trim()) { apply([], ''); dlg.close(); return; }
    const p = parse(raw), A = cfg.actions || {};
    if (p.cmd === 'insight') { renderInsights(true); dlg.close(); panel.scrollIntoView({behavior: 'smooth', block: 'center'}); return; }
    if (p.cmd === 'clear') { setChecked(cfg.table.rows().nodes().toArray(), false); dlg.close(); return; }
    if (p.cmd === 'export') { dlg.close(); A.export(); return; }
    if (p.cmd === 'add') { dlg.close(); A.add(); return; }
    apply(p.filters, raw);
    if (p.cmd === 'select') { setChecked(filteredRows(), true); dlg.close(); return; }
    if (p.cmd === 'delete') {
      const rows = filteredRows().map(tr => ({tr, i: cfg.info(tr)}));
      if (!dlg.open) dlg.showModal();
      askDelete(rows, !p.filters.length); return;
    }
    dlg.close();
  }

  /* ---------- gợi ý phân tích (hiệu ứng gõ chữ) ---------- */
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  function renderInsights(animate) {
    if (!cfg || !cfg.insights) return;
    const items = cfg.insights(), my = ++token;
    const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    panel.hidden = false; list.innerHTML = '';
    if (!items.length) { list.innerHTML = '<li><span>Dữ liệu ổn – chưa có điều gì đáng chú ý.</span></li>'; return; }
    (async () => {
      for (const it of items) {
        if (my !== token) return;
        const li = document.createElement('li'), sp = document.createElement('span');
        li.appendChild(sp); list.appendChild(li);
        if (animate && !reduce) {
          sp.className = 'caret';
          for (let k = 1; k <= it.text.length; k++) { if (my !== token) return; sp.textContent = it.text.slice(0, k); await sleep(9); }
          sp.className = '';
        } else sp.textContent = it.text;
        if (it.q) {
          const b = document.createElement('button'); b.className = 'ai-see'; b.textContent = 'Xem';
          b.onclick = () => execute(it.q); li.appendChild(b);
        }
      }
    })();
  }
  qs('#aiRefresh').onclick = () => renderInsights(true);

  /* ---------- mở / đóng ---------- */
  function open(prefill) {
    if (!cfg) return;
    dlg.showModal(); inp.value = prefill || ''; preview(inp.value); inp.focus();
  }
  trigger.onclick = () => open();
  inp.addEventListener('input', () => preview(inp.value));
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); execute(inp.value); } });
  dlg.addEventListener('click', e => { if (e.target === dlg) dlg.close(); });
  document.addEventListener('keydown', e => {
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test((e.target.tagName || '')) || e.target.isContentEditable;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); dlg.open ? dlg.close() : open(); }
    else if (e.key === '/' && !typing && !dlg.open && cfg) { e.preventDefault(); open(); }
  });

  window.AI = {
    register(c) {
      cfg = c; trigger.hidden = false;
      sugg.innerHTML = '';
      (c.examples || []).forEach(x => {
        const b = document.createElement('button'); b.type = 'button'; b.textContent = x;
        b.onclick = () => { inp.value = x; preview(x); inp.focus(); }; sugg.appendChild(b);
      });
      c.table.on('draw.dt', renderActive);
      renderInsights(true);
    },
    refresh() { if (cfg) { renderInsights(false); renderActive(); } },
    run: execute,
    _parse: parse
  };
})();

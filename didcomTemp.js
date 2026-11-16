/*
 * 温度モニター MyGeotab Add-In
 * Didcom TPM+ で設定した BLE 温度センサー（診断項目 "Remote probe 1〜4 temperature"）の
 * StatusData を MyGeotab API から読み、車両一覧・温度推移・逸脱履歴を表示する。
 * MyGeotab 外（ローカルでファイルを開いた場合）はデモデータで動作する。
 */
/* global geotab */
var geotab = window.geotab || { addin: {} };
geotab.addin = geotab.addin || {};

geotab.addin.didcomTemp = function () {
  'use strict';

  var PROFILES = {
    frozen: { min: -25, max: -18 },
    chilled: { min: 0, max: 10 },
    constant: { min: 15, max: 25 }
  };
  var DEFAULTS = { profile: 'frozen', min: -25, max: -18, minExcursionMin: 5, staleMinutes: 30, diagnostics: [], doorDiagnostics: [], doorInvert: false };
  var DOOR_COLOR = 'rgba(234,179,8,.28)', DOOR_LOOKBACK_MS = 15 * 60000; // 逸脱開始の15分前までのドア開を関連ありとみなす
  var COLORS = ['#2563eb', '#ea580c', '#16a34a', '#9333ea'];
  var AUTO_MS = 5 * 60 * 1000;
  var VERSION = '0.2.2';

  var api = null, demo = false, storageKey = 'didcomTemp:demo';
  var settings = Object.assign({}, DEFAULTS);
  var data = null;            // { vehicles, diagnostics, series: {devId: {diagId: [{t, v}]}} }
  var selectedId = null, timer = null, loading = false;
  var root, $ = function (id) { return document.getElementById(id); };

  // ---------- ユーティリティ ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmtT(d) { return d ? d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '-'; }
  function fmtDur(ms) { var m = Math.round(ms / 60000); return m < 60 ? m + '分' : Math.floor(m / 60) + '時間' + (m % 60) + '分'; }
  function fmtV(v) { return v == null ? '-' : v.toFixed(1) + '℃'; }
  function ago(d) { if (!d) return '-'; var m = Math.round((Date.now() - d) / 60000); return m < 1 ? 'たった今' : m < 60 ? m + '分前' : m < 1440 ? Math.floor(m / 60) + '時間前' : Math.floor(m / 1440) + '日前'; }
  // 「Remote probe 1 temperature」/「リモートプローブ1温度」など英語・日本語どちらの名称にも対応
  var PROBE_RE = /(remote\s*probe|リモート\s*プローブ|プローブ)\s*#?\s*([1-4])/i;
  function isProbeTemp(name) { return PROBE_RE.test(name || '') && /(temperature|温度)/i.test(name || '') && !/(error|alarm|open|short|エラー|アラーム|断線|短絡)/i.test(name || ''); }
  function sensorLabel(name) {
    var m = PROBE_RE.exec(name || '') || /RemoteProbe([1-4])TemperatureId/.exec(name || '');
    return m ? '温度' + m[m.length - 1] : name;
  }
  function call(method, params) { return new Promise(function (res, rej) { api.call(method, params, res, rej); }); }
  function multi(calls) { return new Promise(function (res, rej) { api.multiCall(calls, res, rej); }); }
  function showMsg(text, info) { var el = $('dt-message'); el.textContent = text; el.className = 'dt-message' + (info ? ' info' : ''); el.hidden = !text; }

  function loadSettings() {
    try { var s = JSON.parse(localStorage.getItem(storageKey)); if (s) settings = Object.assign({}, DEFAULTS, s); } catch (e) { /* ignore */ }
  }
  function saveSettings() { try { localStorage.setItem(storageKey, JSON.stringify(settings)); } catch (e) { /* ignore */ } }

  // ---------- データ取得（ライブ） ----------
  function findDiagnostics(query) {
    var names = query ? ['%' + query + '%'] : ['%プローブ%', '%probe%'];
    return multi(names.map(function (n) { return ['Get', { typeName: 'Diagnostic', search: { name: n }, resultsLimit: 200 }]; }))
      .then(function (results) {
        var seen = {}, out = [];
        results.forEach(function (list) { list.forEach(function (d) { if (!seen[d.id]) { seen[d.id] = 1; out.push({ id: d.id, name: d.name }); } }); });
        return out.sort(function (a, b) { return a.name.localeCompare(b.name); });
      });
  }

  // GeotabのKnownId（言語に依存しない固定ID）で検出。失敗時は名称検索にフォールバック。
  var PROBE_IDS = ['RemoteProbe1TemperatureId', 'RemoteProbe2TemperatureId', 'RemoteProbe3TemperatureId', 'RemoteProbe4TemperatureId'];
  var detectInfo = '';
  function detectProbes() {
    var byId = Promise.all(PROBE_IDS.map(function (id) {
      return call('Get', { typeName: 'Diagnostic', search: { id: id } })
        .then(function (list) { return list && list.length ? { id: list[0].id || id, name: list[0].name || id } : null; },
              function () { return null; });
    })).then(function (r) { return r.filter(Boolean); });
    return byId.then(function (found) {
      if (found.length) { detectInfo = '固定ID ' + found.length + '件'; return found; }
      return findDiagnostics('プローブ').then(function (list) {
        var hit = list.filter(function (d) { return isProbeTemp(d.name); });
        detectInfo = '固定ID 0件 / 名称検索 ' + list.length + '件中 ' + hit.length + '件該当';
        return hit;
      }, function (e) { detectInfo = '固定ID 0件 / 名称検索エラー：' + (e && e.message || e); return []; });
    });
  }

  // TPM+の接点センサー（診断「Contact sensor 1/2」）を検出。固定IDは推定のため、名称検索でも補完。
  var DOOR_IDS = ['ContactSensor1Id', 'ContactSensor2Id'];
  var DOOR_RE = /(contact\s*sensor|コンタクト\s*センサー?|接触\s*センサー?|接点\s*センサー?)\s*#?\s*([12])/i;
  function isDoorDiag(name) { return DOOR_RE.test(name || '') && !/(error|alarm|fault|battery|障害|エラー|アラーム|切断|電池|バッテリ)/i.test(name || ''); }
  function detectDoors() {
    return Promise.all(DOOR_IDS.map(function (id) {
      return call('Get', { typeName: 'Diagnostic', search: { id: id } })
        .then(function (l) { return l && l.length ? { id: l[0].id || id, name: l[0].name || id } : null; }, function () { return null; });
    })).then(function (found) {
      found = found.filter(Boolean);
      if (found.length) return found;
      return multi(['%Contact sensor%', '%コンタクト%', '%接触センサー%', '%接点センサー%'].map(function (n) {
        return ['Get', { typeName: 'Diagnostic', search: { name: n }, resultsLimit: 200 }];
      })).then(function (results) {
        var seen = {}, out = [];
        results.forEach(function (l) { l.forEach(function (d) { if (!seen[d.id] && isDoorDiag(d.name)) { seen[d.id] = 1; out.push({ id: d.id, name: d.name }); } }); });
        return out;
      }, function () { return []; });
    });
  }

  function loadLive(hours) {
    var ready = settings.diagnostics.length ? Promise.resolve() :
      detectProbes().then(function (list) {
        // 初回: リモートプローブ1〜4温度（KnownId: RemoteProbeNTemperatureId）を自動選択
        settings.diagnostics = list;
        if (list.length) saveSettings();
      });
    // ドアセンサーは未設定かつ一度も自動設定していない場合のみ自動検出（利用者が外した選択は尊重）
    if (!(settings.doorDiagnostics || []).length && !settings.doorAutoDone) {
      ready = ready.then(detectDoors).then(function (list) {
        if (list.length) { settings.doorDiagnostics = list; settings.doorAutoDone = true; saveSettings(); }
      });
    }
    return ready.then(function () {
      if (!settings.diagnostics.length) { var e = new Error('温度センサーの診断項目が自動検出できませんでした（' + detectInfo + '）。「設定」から対象の診断項目を検索して選んでください。'); e.info = true; throw e; }
      var now = new Date(), from = new Date(now - hours * 3600e3);
      var calls = [['Get', { typeName: 'Device', search: { fromDate: now.toISOString() }, resultsLimit: 5000 }]];
      settings.diagnostics.forEach(function (d) {
        calls.push(['Get', { typeName: 'StatusData', search: { diagnosticSearch: { id: d.id }, fromDate: from.toISOString(), toDate: now.toISOString() }, resultsLimit: 50000 }]);
      });
      var doorDiags = settings.doorDiagnostics || [], nT = settings.diagnostics.length;
      doorDiags.forEach(function (d) {
        calls.push(['Get', { typeName: 'StatusData', search: { diagnosticSearch: { id: d.id }, fromDate: from.toISOString(), toDate: now.toISOString() }, resultsLimit: 50000 }]);
      });
      return multi(calls).then(function (r) {
        var devices = r[0], series = {}, door = {};
        function collect(target, list, diagId) {
          list.forEach(function (s) {
            var dev = s.device && s.device.id; if (!dev || typeof s.data !== 'number') return;
            target[dev] = target[dev] || {};
            (target[dev][diagId] = target[dev][diagId] || []).push({ t: new Date(s.dateTime), v: s.data });
          });
        }
        settings.diagnostics.forEach(function (d, i) { collect(series, r[i + 1], d.id); });
        doorDiags.forEach(function (d, j) { collect(door, r[nT + 1 + j], d.id); });
        [series, door].forEach(function (obj) {
          Object.keys(obj).forEach(function (dev) { Object.keys(obj[dev]).forEach(function (k) { obj[dev][k].sort(function (a, b) { return a.t - b.t; }); }); });
        });
        var vehicles = devices.filter(function (dv) { return series[dv.id]; }).map(function (dv) { return { id: dv.id, name: dv.name }; });
        return { vehicles: vehicles, diagnostics: settings.diagnostics, series: series, door: door, doorDiags: doorDiags, from: from, to: now };
      });
    });
  }

  // ---------- デモデータ ----------
  function loadDemo(hours) {
    var now = Date.now(), from = now - hours * 3600e3, step = 5 * 60000;
    var diags = [{ id: 'd1', name: 'Remote probe 1 temperature' }, { id: 'd2', name: 'Remote probe 2 temperature' }];
    var names = ['名古屋 800 い 25-01', '名古屋 800 い 25-02', '品川 800 か 10-11', '品川 800 か 10-12', '大阪 800 さ 31-07', '大阪 800 さ 31-08', '福岡 800 た 44-20', '札幌 800 な 52-03'];
    var seed = 7; function rnd() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
    var vehicles = [], series = {}, door = {};
    names.forEach(function (n, vi) {
      var id = 'v' + vi; vehicles.push({ id: id, name: n }); series[id] = {};
      diags.forEach(function (d, di) {
        var arr = [], end = vi === 7 ? now - 3 * 3600e3 : now;
        for (var t = from; t <= end; t += step) {
          var v = -21 + di * 0.6 + (rnd() - 0.5) * 1.6;
          var h = (now - t) / 3600e3;
          if (vi === 2 && h > 5 && h < 5.6) v += 9;                 // 荷積み時の逸脱（35分）
          if (vi === 4 && Math.abs(h - 9) < 0.06) v += 7;           // ドア開閉の一瞬（除外される）
          if (vi === 5 && h < 1.2) v += (1.2 - h) * 8;               // 現在も上昇中
          arr.push({ t: new Date(t), v: Math.round(v * 10) / 10 });
        }
        series[id][d.id] = arr;
      });
      // ドア開閉（[何時間前に開いたか, 開いていた分数(null=開いたまま)]）
      var opens = [[12 + vi * 0.7, 6]];               // 荷下ろし
      if (vi === 2) opens.push([5.62, 25]);           // 荷積み → 逸脱の原因
      if (vi === 4) opens.push([9.05, 4]);
      if (vi === 5) opens.push([1.25, null]);         // 開いたまま → 温度上昇中
      if (vi % 3 === 0) opens.push([3 + vi * 0.1, 8]);
      var dArr = [];
      opens.forEach(function (o) {
        var st = now - o[0] * 3600e3; if (st < from) return;
        dArr.push({ t: new Date(st), v: 1 });
        if (o[1] != null) dArr.push({ t: new Date(st + o[1] * 60000), v: 0 });
      });
      dArr.sort(function (a, b) { return a.t - b.t; });
      door[id] = { door1: dArr };
    });
    return Promise.resolve({ vehicles: vehicles, diagnostics: diags, series: series, door: door, doorDiags: [{ id: 'door1', name: '接点センサー1（ドア）' }], from: new Date(from), to: new Date(now) });
  }

  // ---------- 集計 ----------
  function isOut(v) { return v < settings.min || v > settings.max; }
  function hasDoor() { return !!(data && data.doorDiags && data.doorDiags.length); }

  // ドアが開いていた区間 [{start, end, ongoing}]。値が0以外を「開」とし、doorInvert で反転。
  function doorIntervals(devId) {
    var out = [], s = (data.door || {})[devId] || {};
    Object.keys(s).forEach(function (k) {
      var cur = null;
      s[k].forEach(function (p) {
        var open = (p.v !== 0) !== !!settings.doorInvert;
        if (open && !cur) cur = { start: p.t, end: null };
        else if (!open && cur) { cur.end = p.t; out.push(cur); cur = null; }
      });
      if (cur) { cur.end = data.to; cur.ongoing = true; out.push(cur); }
    });
    return out.sort(function (a, b) { return a.start - b.start; });
  }

  function excursions(devId) {
    var out = [], s = data.series[devId] || {};
    data.diagnostics.forEach(function (d) {
      var arr = s[d.id] || [], cur = null;
      arr.forEach(function (p, i) {
        if (isOut(p.v)) {
          if (!cur) cur = { devId: devId, diag: d, start: p.t, end: null, peak: p.v };
          if (Math.abs(p.v - (settings.min + settings.max) / 2) > Math.abs(cur.peak - (settings.min + settings.max) / 2)) cur.peak = p.v;
        } else if (cur) { cur.end = p.t; out.push(cur); cur = null; }
        if (cur && i === arr.length - 1) { cur.lastT = p.t; out.push(cur); }
      });
    });
    return out.filter(function (e) {
      e.duration = (e.end || e.lastT) - e.start;
      return e.duration >= settings.minExcursionMin * 60000 || !e.end;
    });
  }

  function vehicleState(v) {
    var s = data.series[v.id] || {}, latest = [], last = null;
    data.diagnostics.forEach(function (d) {
      var arr = s[d.id]; if (!arr || !arr.length) return;
      var p = arr[arr.length - 1]; latest.push({ diag: d, v: p.v, t: p.t });
      if (!last || p.t > last) last = p.t;
    });
    var status = !latest.length ? 'none'
      : (Date.now() - last > settings.staleMinutes * 60000) ? 'stale'
      : latest.some(function (l) { return isOut(l.v); }) ? 'ng' : 'ok';
    return { latest: latest, last: last, status: status };
  }
  var STATUS_TEXT = { ok: '正常', ng: '逸脱', stale: '通信途絶', none: 'データなし' };
  var STATUS_ORDER = { ng: 0, stale: 1, ok: 2, none: 3 };

  // ---------- 描画 ----------
  function render() {
    var filter = ($('dt-filter').value || '').trim();
    var rows = data.vehicles.map(function (v) { return Object.assign({ v: v }, vehicleState(v)); })
      .sort(function (a, b) { return STATUS_ORDER[a.status] - STATUS_ORDER[b.status] || a.v.name.localeCompare(b.v.name); });
    var events = [];
    data.vehicles.forEach(function (v) {
      var doors = doorIntervals(v.id);
      excursions(v.id).forEach(function (e) {
        var eEnd = e.end || e.lastT;
        e.door = doors.some(function (iv) { return iv.start < eEnd && iv.end > e.start - DOOR_LOOKBACK_MS; });
        events.push(e);
      });
    });
    events.sort(function (a, b) { return b.start - a.start; });
    data.events = events;

    $('kpi-total').textContent = rows.length;
    ['ok', 'ng', 'stale'].forEach(function (k) { $('kpi-' + k).textContent = rows.filter(function (r) { return r.status === k; }).length; });
    $('kpi-events').textContent = events.length;

    if (!selectedId && rows.length) selectedId = rows[0].v.id;
    $('dt-vehicles').innerHTML = rows.filter(function (r) { return !filter || r.v.name.indexOf(filter) >= 0; }).map(function (r) {
      var temps = r.latest.map(function (l) {
        return '<span class="dt-temp' + (isOut(l.v) ? ' ng' : '') + '"><small>' + esc(sensorLabel(l.diag.name)) + '</small>' + fmtV(l.v) + '</span>';
      }).join('');
      return '<tr data-id="' + esc(r.v.id) + '"' + (r.v.id === selectedId ? ' class="selected"' : '') + '>' +
        '<td><span class="dt-status ' + r.status + '">' + STATUS_TEXT[r.status] + '</span></td>' +
        '<td>' + esc(r.v.name) + '</td><td>' + (temps || '-') + '</td><td title="' + fmtT(r.last) + '">' + ago(r.last) + '</td></tr>';
    }).join('') || '<tr><td colspan="4" class="dt-muted">温度データのある車両がありません</td></tr>';

    var names = {}; data.vehicles.forEach(function (v) { names[v.id] = v.name; });
    $('dt-events').innerHTML = events.map(function (e) {
      return '<tr><td>' + esc(names[e.devId]) + '</td><td>' + esc(sensorLabel(e.diag.name)) + '</td><td>' + fmtT(e.start) + '</td>' +
        '<td>' + (e.end ? fmtT(e.end) : '<span class="dt-ongoing">継続中</span>') + '</td><td>' + fmtDur(e.duration) + '</td><td>' + fmtV(e.peak) + '</td>' +
        '<td>' + (!hasDoor() ? '<span class="dt-muted">未設定</span>' : e.door ? '<span class="dt-door-tag">あり</span>' : '-') + '</td></tr>';
    }).join('') || '<tr><td colspan="7" class="dt-muted">期間内の逸脱はありません</td></tr>';

    renderDetail();
  }

  function renderDetail() {
    var v = data.vehicles.filter(function (x) { return x.id === selectedId; })[0];
    $('dt-detail-title').textContent = v ? '温度推移：' + v.name : '温度推移';
    $('dt-range-label').textContent = '基準 ' + settings.min + '〜' + settings.max + '℃';
    var chart = $('dt-chart');
    if (!v) { chart.innerHTML = '<div class="empty">車両を選択してください</div>'; $('dt-legend').innerHTML = ''; $('dt-stats').innerHTML = ''; return; }

    var s = data.series[v.id] || {}, lines = [], all = [];
    data.diagnostics.forEach(function (d, i) { var a = s[d.id]; if (a && a.length) { lines.push({ d: d, a: a, c: COLORS[i % COLORS.length] }); a.forEach(function (p) { all.push(p.v); }); } });
    if (!all.length) { chart.innerHTML = '<div class="empty">期間内のデータがありません</div>'; return; }

    // 表示幅に合わせて描画（縮小で文字が小さくならないように）
    var W = Math.max(360, Math.round(chart.clientWidth || 800)), H = 300, L = 44, R = 12, T = 12, B = 30;
    var lo = Math.min.apply(null, all.concat(settings.min)) - 2, hi = Math.max.apply(null, all.concat(settings.max)) + 2;
    var t0 = +data.from, t1 = +data.to;
    var x = function (t) { return L + (t - t0) / (t1 - t0) * (W - L - R); };
    var y = function (val) { return T + (hi - val) / (hi - lo) * (H - T - B); };
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="温度推移グラフ">';
    svg += '<rect x="' + L + '" y="' + y(settings.max) + '" width="' + (W - L - R) + '" height="' + (y(settings.min) - y(settings.max)) + '" fill="rgba(21,128,61,.08)"/>';
    // ドア開の区間を背景に帯で表示（マウスを乗せると時刻を表示）
    var doors = hasDoor() ? doorIntervals(v.id) : [], doorMs = 0;
    doors.forEach(function (iv) {
      var s0 = Math.max(+iv.start, t0), s1 = Math.min(+iv.end, t1); if (s1 <= s0) return;
      doorMs += iv.end - iv.start;
      var w = Math.max(2, x(s1) - x(s0));
      svg += '<rect x="' + x(s0).toFixed(1) + '" y="' + T + '" width="' + w.toFixed(1) + '" height="' + (H - T - B) + '" fill="' + DOOR_COLOR + '">' +
        '<title>ドア開 ' + fmtT(iv.start) + '〜' + (iv.ongoing ? '開いたまま' : fmtT(iv.end)) + '（' + fmtDur(iv.end - iv.start) + '）</title></rect>';
    });
    var stepV = (hi - lo) > 20 ? 5 : 2;
    for (var g = Math.ceil(lo / stepV) * stepV; g <= hi; g += stepV) {
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(g) + '" y2="' + y(g) + '" stroke="#e5e7eb"/>' +
        '<text x="' + (L - 6) + '" y="' + (y(g) + 4) + '" font-size="11" text-anchor="end" fill="#6b7280">' + g + '</text>';
    }
    [settings.min, settings.max].forEach(function (th) {
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(th) + '" y2="' + y(th) + '" stroke="#dc2626" stroke-dasharray="4 3"/>';
    });
    var ticks = Math.max(2, Math.min(6, Math.floor((W - L - R) / 110)));
    for (var k = 0; k <= ticks; k++) {
      var tt = t0 + (t1 - t0) * k / ticks;
      svg += '<text x="' + x(tt) + '" y="' + (H - 8) + '" font-size="11" text-anchor="middle" fill="#6b7280">' + fmtT(new Date(tt)) + '</text>';
    }
    lines.forEach(function (ln) {
      var pts = ln.a.map(function (p) { return x(+p.t).toFixed(1) + ',' + y(p.v).toFixed(1); }).join(' ');
      svg += '<polyline points="' + pts + '" fill="none" stroke="' + ln.c + '" stroke-width="1.8"/>';
    });
    chart.innerHTML = svg + '</svg>';
    $('dt-legend').innerHTML = lines.map(function (ln) { return '<span><i style="background:' + ln.c + '"></i>' + esc(sensorLabel(ln.d.name)) + '</span>'; }).join('') +
      '<span><i style="background:#dc2626"></i>しきい値</span>' +
      (hasDoor() ? '<span><i class="dt-door-swatch" style="background:' + DOOR_COLOR + '"></i>ドア開</span>' : '');

    var avg = all.reduce(function (a, b) { return a + b; }, 0) / all.length;
    var exMs = data.events.filter(function (e) { return e.devId === v.id; }).reduce(function (a, e) { return a + e.duration; }, 0);
    $('dt-stats').innerHTML = [['最低', fmtV(Math.min.apply(null, all))], ['最高', fmtV(Math.max.apply(null, all))], ['平均', fmtV(avg)], ['逸脱時間合計', exMs ? fmtDur(exMs) : 'なし']]
      .concat(hasDoor() ? [['ドア開 回数', doors.length + '回'], ['ドア開 合計', doorMs ? fmtDur(doorMs) : 'なし']] : [])
      .map(function (p) { return '<div><span>' + p[0] + '</span><strong>' + p[1] + '</strong></div>'; }).join('');
  }

  // ---------- 更新・CSV ----------
  function refresh() {
    if (loading) return; loading = true;
    $('dt-refresh').disabled = true;
    var hours = +$('dt-hours').value;
    (demo ? loadDemo(hours) : loadLive(hours)).then(function (d) {
      data = d;
      showMsg(demo ? 'MyGeotab外で開いているため、デモデータを表示しています。'
        : !d.vehicles.length ? '監視対象：' + d.diagnostics.map(function (x) { return x.name; }).join('、') + '。期間内にデータがある車両はまだありません（センサーの有効化・MAC登録後に表示されます）。'
        : '', true);
      render();
    }).catch(function (e) {
      if (e && e.info) showMsg(e.message, true);
      else showMsg('データ取得に失敗しました：' + (e && e.message ? e.message : e));
    }).then(function () { loading = false; $('dt-refresh').disabled = false; });
  }
  function schedule() {
    clearInterval(timer); timer = null;
    if ($('dt-auto').checked) timer = setInterval(refresh, AUTO_MS);
  }

  function exportCsv() {
    if (!data || !data.events) return;
    var names = {}; data.vehicles.forEach(function (v) { names[v.id] = v.name; });
    var q = function (s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
    var lines = [['車両', 'センサー', '開始', '終了', '継続時間(分)', '最高/最低(℃)', '下限(℃)', '上限(℃)', 'ドア開'].map(q).join(',')];
    data.events.forEach(function (e) {
      lines.push([names[e.devId], sensorLabel(e.diag.name), fmtT(e.start), e.end ? fmtT(e.end) : '継続中', Math.round(e.duration / 60000), e.peak, settings.min, settings.max,
        !hasDoor() ? '未設定' : e.door ? 'あり' : 'なし'].map(q).join(','));
    });
    var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }); // BOM付き: Excelで文字化けしない
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = '温度逸脱履歴_' + new Date().toISOString().slice(0, 10) + '.csv'; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // ---------- 設定画面 ----------
  // 診断項目の検索・選択リスト（温度センサー用とドアセンサー用で共用）
  function Picker(prefix, defaultQuery) { this.p = prefix; this.q = defaultQuery; this.draft = []; }
  Picker.prototype.render = function (list, fromSearch) {
    var chosen = {}; this.draft.forEach(function (d) { chosen[d.id] = 1; });
    $(this.p + '-list').innerHTML = list.map(function (d) {
      return '<label><input type="checkbox" data-id="' + esc(d.id) + '" data-name="' + esc(d.name) + '"' + (chosen[d.id] ? ' checked' : '') + '> ' +
        esc(d.name) + ' <code>' + esc(d.id) + '</code></label>';
    }).join('') || '<p class="dt-muted" style="padding:8px">' + (fromSearch ? '該当なし（別のキーワードで検索してください）' : '未選択（検索して選んでください）') + '</p>';
  };
  // 表示中のチェック状態を反映した選択結果（表示されていない選択済み項目は保持）
  Picker.prototype.collect = function () {
    var shown = {}, picked = [];
    Array.prototype.forEach.call($(this.p + '-list').querySelectorAll('input[type=checkbox]'), function (cb) {
      shown[cb.dataset.id] = 1; if (cb.checked) picked.push({ id: cb.dataset.id, name: cb.dataset.name });
    });
    return this.draft.filter(function (d) { return !shown[d.id]; }).concat(picked);
  };
  Picker.prototype.search = function () {
    var self = this;
    self.draft = self.collect();
    if (demo) { self.render(self.draft); return; }
    $(self.p + '-list').innerHTML = '<p class="dt-muted" style="padding:8px">検索中…</p>';
    findDiagnostics($(self.p + '-query').value.trim() || self.q).then(function (l) { self.render(l, true); }).catch(function (e) {
      $(self.p + '-list').innerHTML = '<p class="dt-muted" style="padding:8px">検索に失敗しました：' + esc(e && e.message || e) + '</p>';
    });
  };
  var tempPicker = new Picker('set-diag', ''), doorPicker = new Picker('set-door', 'Contact sensor');

  function openSettings() {
    $('set-profile').value = settings.profile; $('set-min').value = settings.min; $('set-max').value = settings.max;
    $('set-minex').value = settings.minExcursionMin; $('set-stale').value = settings.staleMinutes;
    $('set-door-invert').checked = !!settings.doorInvert;
    tempPicker.draft = (settings.diagnostics.length ? settings.diagnostics : (data ? data.diagnostics : [])).slice();
    doorPicker.draft = ((settings.doorDiagnostics || []).length ? settings.doorDiagnostics : (demo && data ? data.doorDiags : [])).slice();
    tempPicker.render(tempPicker.draft); doorPicker.render(doorPicker.draft);
    $('dt-settings').hidden = false;
  }
  function saveFromForm() {
    settings.profile = $('set-profile').value;
    settings.min = parseFloat($('set-min').value); settings.max = parseFloat($('set-max').value);
    settings.minExcursionMin = Math.max(0, parseInt($('set-minex').value, 10) || 0);
    settings.staleMinutes = Math.max(1, parseInt($('set-stale').value, 10) || 30);
    if (!(settings.min < settings.max)) { alert('下限は上限より小さい値にしてください。'); return; }
    settings.doorInvert = $('set-door-invert').checked;
    if (!demo) {
      settings.diagnostics = tempPicker.collect();
      settings.doorDiagnostics = doorPicker.collect();
      settings.doorAutoDone = true;
    }
    saveSettings(); $('dt-settings').hidden = true; refresh();
  }

  function bind() {
    $('dt-refresh').onclick = refresh;
    var rz = null;
    window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (data) renderDetail(); }, 150); });
    $('dt-hours').onchange = refresh;
    $('dt-auto').onchange = schedule;
    $('dt-filter').oninput = function () { if (data) render(); };
    $('dt-csv').onclick = exportCsv;
    $('dt-vehicles').onclick = function (e) {
      var tr = e.target.closest('tr[data-id]'); if (!tr) return;
      selectedId = tr.dataset.id; render();
    };
    $('dt-open-settings').onclick = openSettings;
    $('set-cancel').onclick = function () { $('dt-settings').hidden = true; };
    $('set-save').onclick = saveFromForm;
    $('set-diag-search').onclick = function () { tempPicker.search(); };
    $('set-door-search').onclick = function () { doorPicker.search(); };
    $('set-profile').onchange = function () {
      var p = PROFILES[this.value]; if (p) { $('set-min').value = p.min; $('set-max').value = p.max; }
    };
  }

  function start(a, isDemo) {
    api = a; demo = isDemo; root = $('didcomTemp');
    var badge = $('dt-source'); badge.textContent = (demo ? 'デモ' : 'ライブ') + ' v' + VERSION; badge.className = 'dt-badge' + (demo ? ' demo' : '');
    var ready = demo ? Promise.resolve() : new Promise(function (res) {
      api.getSession(function (s) { storageKey = 'didcomTemp:' + s.database; res(); });
    });
    return ready.then(function () { loadSettings(); bind(); });
  }

  return {
    initialize: function (a, state, callback) { start(a, false).then(function () { callback(); }, function () { callback(); }); },
    focus: function (a) { api = a; refresh(); schedule(); },
    blur: function () { clearInterval(timer); timer = null; },
    _standalone: function () { start(null, true).then(function () { refresh(); schedule(); }); }
  };
};

// MyGeotab外（ローカルでHTMLを開いた場合）はデモモードで起動。
// MyGeotab内では initialize が呼ばれるので、2秒待っても呼ばれなければデモとみなす。
(function () {
  var orig = geotab.addin.didcomTemp;
  geotab.addin.didcomTemp = function () {
    var inst = orig();
    var init = inst.initialize;
    inst.initialize = function () { window.__didcomTempInit = true; return init.apply(inst, arguments); };
    return inst;
  };
  window.addEventListener('load', function () {
    setTimeout(function () { if (!window.__didcomTempInit) orig()._standalone(); }, 2000);
  });
})();

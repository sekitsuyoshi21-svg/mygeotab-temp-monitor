/*
 * 庫内温度 — Geotab Drive Add-In（運転手のスマホ・タブレット用）
 * Driveで選んでいる車両の Didcom TPM+ 温度センサー（Remote probe 1〜4 temperature）と
 * 接点センサー（ドア）の最新値・状態・直近6時間のグラフを表示する。
 * Drive外（ローカルでHTMLを開いた場合）は2秒後にデモデータで表示する。
 */
/* global geotab */
var geotab = window.geotab || { addin: {} };
geotab.addin = geotab.addin || {};

geotab.addin.didcomDrive = function () {
  'use strict';

  var VERSION = '0.1.0';
  var PROFILES = {
    frozen: { min: -25, max: -18, label: '冷凍 -25〜-18℃' },
    chilled: { min: 0, max: 10, label: '冷蔵 0〜10℃' },
    constant: { min: 15, max: 25, label: '定温 15〜25℃' }
  };
  var HOURS = 6, STALE_MIN = 60, RISE_RATE = 3, RISE_WIN_MIN = 60, AUTO_MS = 2 * 60000;
  var COLORS = ['#2563eb', '#ea580c', '#16a34a', '#9333ea'];
  var PROBE_IDS = ['RemoteProbe1TemperatureId', 'RemoteProbe2TemperatureId', 'RemoteProbe3TemperatureId', 'RemoteProbe4TemperatureId'];
  var PROBE_RE = /(remote\s*probe|リモート\s*プローブ|プローブ)\s*#?\s*([1-4])/i;
  var DOOR_IDS = ['ContactSensor1Id', 'ContactSensor2Id'];
  var DOOR_RE = /(contact\s*sensor|コンタクト\s*センサー?|接触\s*センサー?|接点\s*センサー?)\s*#?\s*([12])/i;

  var api = null, demo = false, driveState = null, storageKey = 'didcomDrive:demo';
  var settings = { profile: 'frozen', doorInvert: false, deviceId: null, diagnostics: null, doors: null };
  var timer = null, loading = false;
  var $ = function (id) { return document.getElementById(id); };

  // ---------- ユーティリティ ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function hm(d) { return d.toLocaleTimeString('ja-JP', { timeZone: 'Asia/Tokyo', hour: '2-digit', minute: '2-digit' }); }
  function ago(d) { if (!d) return '-'; var m = Math.round((Date.now() - d) / 60000); return m < 1 ? 'たった今' : m < 60 ? m + '分前' : m < 1440 ? Math.floor(m / 60) + '時間前' : Math.floor(m / 1440) + '日前'; }
  function fmtRate(r) { return r == null ? '' : (r >= 0 ? '↑+' : '↓') + r.toFixed(1) + '℃/時'; }
  function label(name, i) { var m = PROBE_RE.exec(name || '') || /RemoteProbe([1-4])/.exec(name || ''); return '温度' + (m ? m[m.length - 1] : i + 1); }
  function call(method, params) { return new Promise(function (res, rej) { api.call(method, params, res, rej); }); }
  function multi(calls) { return new Promise(function (res, rej) { api.multiCall(calls, res, rej); }); }
  function msg(text, info) { var el = $('dd-message'); el.textContent = text || ''; el.className = 'dd-message' + (info ? ' info' : ''); el.hidden = !text; }
  function prof() { return PROFILES[settings.profile] || PROFILES.frozen; }
  function load() { try { var s = JSON.parse(localStorage.getItem(storageKey)); if (s) Object.assign(settings, s); } catch (e) { /* ignore */ } }
  function save() { try { localStorage.setItem(storageKey, JSON.stringify(settings)); } catch (e) { /* ignore */ } }

  // ---------- 診断項目の検出（管理画面の温度モニターと同じ方法） ----------
  function byKnownIds(ids) {
    return Promise.all(ids.map(function (id) {
      return call('Get', { typeName: 'Diagnostic', search: { id: id } })
        .then(function (l) { return l && l.length ? { id: l[0].id || id, name: l[0].name || id } : null; }, function () { return null; });
    })).then(function (r) { return r.filter(Boolean); });
  }
  function byName(patterns, test) {
    return multi(patterns.map(function (n) { return ['Get', { typeName: 'Diagnostic', search: { name: n }, resultsLimit: 200 }]; }))
      .then(function (results) {
        var seen = {}, out = [];
        results.forEach(function (l) { l.forEach(function (d) { if (!seen[d.id] && test(d.name)) { seen[d.id] = 1; out.push({ id: d.id, name: d.name }); } }); });
        return out.sort(function (a, b) { return a.name.localeCompare(b.name); });
      }, function () { return []; });
  }
  function detect() {
    if (settings.diagnostics && settings.diagnostics.length && settings.doors) return Promise.resolve();
    var probes = byKnownIds(PROBE_IDS).then(function (f) {
      return f.length ? f : byName(['%プローブ%', '%probe%'], function (n) {
        return PROBE_RE.test(n || '') && /(temperature|温度)/i.test(n || '') && !/(error|alarm|open|short|エラー|アラーム|断線|短絡)/i.test(n || '');
      });
    });
    var doors = byKnownIds(DOOR_IDS).then(function (f) {
      return f.length ? f : byName(['%Contact sensor%', '%コンタクト%', '%接触センサー%', '%接点センサー%'], function (n) {
        return DOOR_RE.test(n || '') && !/(error|alarm|fault|battery|障害|エラー|アラーム|切断|電池|バッテリ)/i.test(n || '');
      });
    });
    return Promise.all([probes, doors]).then(function (r) {
      settings.diagnostics = r[0]; settings.doors = r[1];
      if (r[0].length) save();
    });
  }

  // ---------- 車両の決定 ----------
  // Driveで選択中の車両（state.device）を優先。無ければ一覧から選んでもらう。
  function stateDeviceId() {
    var d = driveState && driveState.device;
    return d && (d.id || (typeof d === 'string' ? d : null)) || null;
  }
  function resolveDevice() {
    var fromDrive = stateDeviceId();
    if (fromDrive) { $('dd-pick-wrap').hidden = true; return call('Get', { typeName: 'Device', search: { id: fromDrive } }).then(function (l) { return l[0] || { id: fromDrive, name: fromDrive }; }); }
    return call('Get', { typeName: 'Device', search: { fromDate: new Date().toISOString() }, resultsLimit: 5000 }).then(function (list) {
      list.sort(function (a, b) { return a.name.localeCompare(b.name, 'ja'); });
      var sel = $('dd-pick');
      sel.innerHTML = '<option value="">（車両を選択）</option>' + list.map(function (d) { return '<option value="' + esc(d.id) + '">' + esc(d.name) + '</option>'; }).join('');
      $('dd-pick-wrap').hidden = false;
      var cur = list.filter(function (d) { return d.id === settings.deviceId; })[0];
      sel.value = cur ? cur.id : '';
      if (!cur) { $('dd-settings').open = true; return null; }
      return cur;
    });
  }

  // ---------- データ取得 ----------
  function loadLive() {
    return detect().then(function () {
      if (!settings.diagnostics.length) throw Object.assign(new Error('温度センサーの診断項目が見つかりませんでした。管理者にお問い合わせください。'), { info: true });
      return resolveDevice();
    }).then(function (dev) {
      if (!dev) return null;
      var now = new Date(), from = new Date(now - HOURS * 3600e3);
      // ドアは開いたまま期間の前から続いている場合があるので、少し前から読む
      var dFrom = new Date(now - (HOURS + 24) * 3600e3);
      var calls = settings.diagnostics.map(function (d) {
        return ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: dev.id }, diagnosticSearch: { id: d.id }, fromDate: from.toISOString(), toDate: now.toISOString() } }];
      }).concat(settings.doors.map(function (d) {
        return ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: dev.id }, diagnosticSearch: { id: d.id }, fromDate: dFrom.toISOString(), toDate: now.toISOString() } }];
      }));
      return multi(calls).then(function (r) {
        var n = settings.diagnostics.length;
        function pts(list) { return (list || []).filter(function (s) { return typeof s.data === 'number'; }).map(function (s) { return { t: new Date(s.dateTime), v: s.data }; }).sort(function (a, b) { return a.t - b.t; }); }
        return {
          device: dev, from: from, to: now,
          temps: settings.diagnostics.map(function (d, i) { return { diag: d, label: label(d.name, i), pts: pts(r[i]) }; }),
          doors: settings.doors.map(function (d, j) { return pts(r[n + j]); })
        };
      });
    });
  }

  function loadDemo() {
    var now = Date.now(), from = now - HOURS * 3600e3, step = 30 * 60000, seed = 3;
    function rnd() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
    var temps = [0, 1].map(function (i) {
      var pts = [];
      for (var t = now - 4 * 60000 - Math.floor((HOURS * 3600e3 - 4 * 60000) / step) * step; t <= now - 4 * 60000; t += step) {
        var h = (now - t) / 3600e3, v = -21 + i * 0.7 + (rnd() - 0.5) * 1.2;
        if (h < 1.1) v += (1.1 - h) * 3.8;                 // 直近1時間で上昇中（急上昇の例）
        if (Math.abs(h - 3.5) < 0.3) v += 4;              // 荷下ろしでの一時上昇
        pts.push({ t: new Date(t), v: Math.round(v * 10) / 10 });
      }
      return { diag: { id: 'p' + i, name: 'Remote probe ' + (i + 1) + ' temperature' }, label: '温度' + (i + 1), pts: pts };
    });
    settings.doors = [{ id: 'door1', name: '接点センサー1' }];
    var doors = [[{ t: new Date(now - 3.6 * 3600e3), v: 1 }, { t: new Date(now - 3.45 * 3600e3), v: 0 }, { t: new Date(now - 12 * 60000), v: 1 }]];
    return Promise.resolve({ device: { id: 'demo', name: '名古屋 800 い 25-56' }, from: new Date(from), to: new Date(now), temps: temps, doors: doors });
  }

  // ---------- 集計 ----------
  function rateAt(pts) {
    if (pts.length < 2) return null;
    var last = pts[pts.length - 1], win = RISE_WIN_MIN * 60000, base = null;
    for (var i = pts.length - 2; i >= 0; i--) { if (last.t - pts[i].t >= win) { base = pts[i]; break; } }
    if (!base) base = pts[0];
    var dt = last.t - base.t;
    if (dt < 10 * 60000 || dt > 2 * win) return null;
    return (last.v - base.v) / dt * 3600e3;
  }
  function isOpen(v) { return (v !== 0) !== !!settings.doorInvert; }
  function doorIntervals(arr, to) {
    var out = [], cur = null;
    arr.forEach(function (p) {
      if (isOpen(p.v) && !cur) cur = { start: p.t };
      else if (!isOpen(p.v) && cur) { cur.end = p.t; out.push(cur); cur = null; }
    });
    if (cur) { cur.end = to; cur.ongoing = true; out.push(cur); }
    return out;
  }

  // ---------- 描画 ----------
  function render(d) {
    var p = prof();
    $('dd-vehicle').textContent = d.device.name;
    $('dd-profile-label').textContent = '基準 ' + p.label;
    var last = null, anyOut = false, anyRise = false, outText = [];
    var html = d.temps.map(function (s, i) {
      var lp = s.pts[s.pts.length - 1];
      if (!lp) return '<div class="dd-sensor"><span class="dd-label"><i class="dd-dot" style="background:' + COLORS[i % 4] + '"></i>' + esc(s.label) + '</span><strong>-</strong><small>データなし</small></div>';
      if (!last || lp.t > last) last = lp.t;
      var r = rateAt(s.pts), out = lp.v < p.min || lp.v > p.max, rise = !out && r != null && r >= RISE_RATE;
      if (out) { anyOut = true; outText.push(s.label + (lp.v > p.max ? ' 高い' : ' 低い')); }
      if (rise) anyRise = true;
      return '<div class="dd-sensor' + (out ? ' ng' : rise ? ' rise' : '') + '"><span class="dd-label"><i class="dd-dot" style="background:' + COLORS[i % 4] + '"></i>' + esc(s.label) + '</span>' +
        '<strong>' + lp.v.toFixed(1) + '<span style="font-size:18px">℃</span></strong><small>' + (fmtRate(r) || '&nbsp;') + '</small></div>';
    }).join('');
    $('dd-sensors').innerHTML = html;

    var stale = !last || Date.now() - last > STALE_MIN * 60000;
    var st = !last ? 'none' : stale ? 'stale' : anyOut ? 'ng' : anyRise ? 'rise' : 'ok';
    var TEXT = { ok: '正常', ng: '温度逸脱', rise: '急上昇', stale: 'データ未受信', none: 'データなし' };
    var NOTE = {
      ok: '基準温度の範囲内です',
      ng: outText.join('・') + '｜冷凍機とドアを確認してください',
      rise: '温度が上がっています｜ドアと冷凍機を確認してください',
      stale: '最後のデータは ' + (last ? ago(last) : '-') + '｜センサーと通信を確認してください',
      none: '直近' + HOURS + '時間のデータがありません'
    };
    $('dd-status').className = 'dd-status ' + st;
    $('dd-status-text').textContent = TEXT[st];
    $('dd-status-note').textContent = NOTE[st];

    // ドア
    var doorEl = $('dd-door');
    if (!settings.doors || !settings.doors.length) { doorEl.className = 'dd-tile'; doorEl.querySelector('strong').textContent = '-'; doorEl.querySelector('small').textContent = 'センサーなし'; }
    else {
      var lastDoor = null;
      d.doors.forEach(function (arr) { var x = arr[arr.length - 1]; if (x && (!lastDoor || x.t > lastDoor.t)) lastDoor = x; });
      var open = d.doors.some(function (arr) { var x = arr[arr.length - 1]; return x && isOpen(x.v); });
      doorEl.className = 'dd-tile' + (open ? ' open' : '');
      doorEl.querySelector('strong').textContent = lastDoor ? (open ? '開' : '閉') : '-';
      doorEl.querySelector('small').textContent = lastDoor ? hm(lastDoor.t) + (open ? ' から開' : ' に閉') : 'データなし';
    }
    var lastEl = $('dd-last');
    lastEl.className = 'dd-tile' + (stale && last ? ' old' : '');
    lastEl.querySelector('strong').textContent = last ? ago(last) : '-';
    lastEl.querySelector('small').textContent = last ? hm(last) + ' 計測' : '';

    renderChart(d);
  }

  function renderChart(d) {
    var p = prof(), W = 360, H = 170, L = 34, R = 8, T = 8, B = 22, PW = W - L - R, PH = H - T - B;
    var all = []; d.temps.forEach(function (s) { s.pts.forEach(function (x) { all.push(x.v); }); });
    var lo = Math.min.apply(null, all.concat([p.min])) - 1, hi = Math.max.apply(null, all.concat([p.max])) + 1;
    var t0 = +d.from, t1 = +d.to;
    function X(t) { return L + (t - t0) / (t1 - t0) * PW; }
    function Y(v) { return T + (hi - v) / (hi - lo) * PH; }
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none" role="img" aria-label="直近' + HOURS + '時間の温度">';
    // ドア開の帯
    d.doors.forEach(function (arr) {
      doorIntervals(arr, d.to).forEach(function (iv) {
        var a = Math.max(t0, +iv.start), b = Math.min(t1, +iv.end); if (b <= a) return;
        s += '<rect x="' + X(a) + '" y="' + T + '" width="' + Math.max(2, X(b) - X(a)) + '" height="' + PH + '" fill="rgba(234,179,8,.35)"/>';
      });
    });
    // 基準温度の帯
    s += '<rect x="' + L + '" y="' + Y(p.max) + '" width="' + PW + '" height="' + (Y(p.min) - Y(p.max)) + '" fill="rgba(22,163,74,.10)"/>';
    [p.min, p.max].forEach(function (v) { s += '<line x1="' + L + '" x2="' + (L + PW) + '" y1="' + Y(v) + '" y2="' + Y(v) + '" stroke="#16a34a" stroke-dasharray="4 3" stroke-width="1"/>'; });
    // 縦軸
    var step = (hi - lo) > 20 ? 10 : (hi - lo) > 8 ? 5 : 2;
    for (var v = Math.ceil(lo / step) * step; v <= hi; v += step) {
      s += '<line x1="' + L + '" x2="' + (L + PW) + '" y1="' + Y(v) + '" y2="' + Y(v) + '" stroke="#e2e8f0"/>' +
        '<text x="' + (L - 4) + '" y="' + (Y(v) + 4) + '" font-size="11" text-anchor="end" fill="#64748b">' + v + '</text>';
    }
    // 時間軸（1時間ごと、2時間ごとに時刻）
    var hr = 3600e3, first = Math.ceil(t0 / hr) * hr;
    for (var t = first; t <= t1; t += hr) {
      var lbl = new Date(t).toLocaleTimeString('ja-JP', { timeZone: 'Asia/Tokyo', hour: '2-digit', minute: '2-digit' });
      s += '<line x1="' + X(t) + '" x2="' + X(t) + '" y1="' + T + '" y2="' + (T + PH) + '" stroke="#f1f5f9"/>';
      if (Math.round((t - first) / hr) % 2 === 0) s += '<text x="' + X(t) + '" y="' + (H - 6) + '" font-size="11" text-anchor="middle" fill="#64748b">' + lbl + '</text>';
    }
    // 温度の線と点（範囲外の点は赤）
    d.temps.forEach(function (sr, i) {
      if (!sr.pts.length) return;
      s += '<polyline fill="none" stroke="' + COLORS[i % 4] + '" stroke-width="2" stroke-linejoin="round" points="' + sr.pts.map(function (x) { return X(+x.t).toFixed(1) + ',' + Y(x.v).toFixed(1); }).join(' ') + '"/>';
      sr.pts.forEach(function (x) {
        var out = x.v < p.min || x.v > p.max;
        s += '<circle cx="' + X(+x.t).toFixed(1) + '" cy="' + Y(x.v).toFixed(1) + '" r="' + (out ? 3 : 2.2) + '" fill="' + (out ? '#dc2626' : COLORS[i % 4]) + '"/>';
      });
    });
    s += '</svg>';
    $('dd-chart').innerHTML = s;
    $('dd-legend').innerHTML = d.temps.map(function (sr, i) { return '<span><i style="background:' + COLORS[i % 4] + '"></i>' + esc(sr.label) + '</span>'; }).join('') +
      '<span><i class="band"></i>基準</span>' + (settings.doors && settings.doors.length || demo ? '<span><i class="door"></i>ドア開</span>' : '');
  }

  // ---------- 更新 ----------
  function applyDriving() {
    var driving = !!(driveState && driveState.driving);
    $('didcomDrive').classList.toggle('driving', driving);
  }
  function refresh() {
    if (loading) return;
    loading = true; $('dd-refresh').classList.add('spin'); applyDriving();
    (demo ? loadDemo() : loadLive()).then(function (d) {
      if (!d) { msg('表示する車両を「表示設定」で選んでください。', true); $('dd-status').className = 'dd-status none'; $('dd-status-text').textContent = '車両未選択'; $('dd-status-note').textContent = ''; return; }
      msg(''); render(d);
    }).catch(function (e) {
      var t = e && (e.message || e.name) || String(e);
      if (/permission|denied|NotAuthorized|権限/i.test(t)) t = 'データを見る権限がありません。管理者にユーザー権限（ステータスデータの閲覧）の確認を依頼してください。';
      else if (!e.info && /network|offline|timeout|Failed to fetch/i.test(t)) t = '通信できませんでした。電波の良い場所で「更新」を押してください。';
      msg(t, e && e.info);
    }).then(function () { loading = false; $('dd-refresh').classList.remove('spin'); });
  }
  function schedule() { clearInterval(timer); timer = setInterval(refresh, AUTO_MS); }

  function start(a, state, isDemo) {
    api = a; driveState = state || null; demo = isDemo;
    var badge = $('dd-badge'); badge.textContent = (demo ? 'デモ' : 'ライブ') + ' v' + VERSION; badge.className = 'dd-badge' + (demo ? ' demo' : '');
    var ready = demo ? Promise.resolve() : new Promise(function (res) {
      try { api.getSession(function (s) { storageKey = 'didcomDrive:' + s.database; res(); }, function () { res(); }); } catch (e) { res(); }
    });
    return ready.then(function () {
      load();
      $('dd-profile').value = settings.profile;
      $('dd-invert').checked = !!settings.doorInvert;
      $('dd-refresh').onclick = refresh;
      $('dd-profile').onchange = function () { settings.profile = this.value; save(); refresh(); };
      $('dd-invert').onchange = function () { settings.doorInvert = this.checked; save(); refresh(); };
      $('dd-pick').onchange = function () { settings.deviceId = this.value || null; save(); refresh(); };
      document.addEventListener('visibilitychange', function () { if (!document.hidden) refresh(); });
    });
  }

  return {
    initialize: function (a, state, callback) { start(a, state, false).then(function () { callback(); }, function () { callback(); }); },
    focus: function (a, state) { api = a; if (state) driveState = state; refresh(); schedule(); },
    blur: function () { clearInterval(timer); timer = null; },
    _standalone: function () { start(null, { driving: /driving/.test(location.hash) }, true).then(function () { refresh(); schedule(); }); }
  };
};

// Drive外（ローカルでHTMLを開いた場合）はデモで起動。Drive内では initialize が呼ばれる。
(function () {
  var orig = geotab.addin.didcomDrive;
  geotab.addin.didcomDrive = function () {
    var inst = orig(), init = inst.initialize;
    inst.initialize = function () { window.__didcomDriveInit = true; return init.apply(inst, arguments); };
    return inst;
  };
  window.addEventListener('load', function () {
    setTimeout(function () { if (!window.__didcomDriveInit) orig()._standalone(); }, 2000);
  });
})();

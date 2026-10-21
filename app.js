/* 교수법 프로그램 포인트 현황
 * 관리대장 엑셀의 '교수별현황' · '프로그램별현황' 시트(수식 계산 결과)를 읽어 화면에 보여줍니다.
 * 모든 처리는 브라우저 안에서만 이루어지며 데이터는 저장·전송되지 않습니다.
 */
const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const nf = n => Number(n || 0).toLocaleString('ko-KR');

let S = null; // 업로드된 데이터

/* ---------- 엑셀 읽기 ---------- */
// 교수별현황: 3행=프로그램별 상한, 4행=머리글, 5행~=데이터 (A 순위, B 소속, C 구분, D 교번, E 성명,
//             F 총 인정, G 적립 합계, H 상한 초과분, I 참여 횟수, J~Q 프로그램별 인정 포인트)
// 프로그램별현황: 5행~ (A 프로그램, B 상한, C 참여 횟수, D 참여 인원, E 인정 합계, F 초과분, G 상한 도달, H~J 구분별 인원)
function parse(wb) {
  const ws = wb.Sheets['교수별현황'], pg = wb.Sheets['프로그램별현황'];
  if (!ws || !pg) throw new Error('이 관리대장 양식이 아닙니다. 교수별현황 시트와 프로그램별현황 시트가 필요합니다.');

  const a = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const progs = [];
  for (let j = 9; j < 17; j++) {
    if (String(a[3][j]).trim()) progs.push({ col: j, name: String(a[3][j]), cap: typeof a[2][j] === 'number' ? a[2][j] : null });
  }
  const profs = a.slice(4)
    .filter(r => String(r[3]).trim() !== '')
    .map(r => ({
      rank: r[0], dept: r[1], grp: r[2], id: String(r[3]).trim(), name: r[4],
      total: +r[5] || 0, raw: +r[6] || 0, over: +r[7] || 0, cnt: +r[8] || 0,
      pts: progs.map(p => +r[p.col] || 0)
    }));

  const b = XLSX.utils.sheet_to_json(pg, { header: 1, defval: '' });
  const gn = [7, 8, 9].map(i => (String(b[3][i]).match(/\((.*)\)/) || [0, ''])[1]);
  const pr = b.slice(4)
    .filter(r => String(r[0]).trim() !== '' && r[0] !== '합계' && !String(r[0]).startsWith('※'))
    .map(r => ({
      name: r[0], cap: r[1], cnt: +r[2] || 0, ppl: +r[3] || 0, pts: +r[4] || 0,
      over: +r[5] || 0, hit: r[6], g: [+r[7] || 0, +r[8] || 0, +r[9] || 0]
    }));

  if (!profs.length) throw new Error('집계 값이 비어 있습니다. 엑셀에서 파일을 열어 저장한 뒤 다시 올려 주세요.');
  return { progs, profs, pr, gn };
}

/* ---------- 관리자 현황 ---------- */
function render() {
  const P = S.profs, act = P.filter(p => p.cnt > 0), sum = act.reduce((s, p) => s + p.total, 0);
  $('kpis').innerHTML = [
    ['등록 교수', nf(P.length) + '명'],
    ['참여 교수', nf(act.length) + '명'],
    ['총 인정 포인트', nf(sum)],
    ['참여자 평균', act.length ? (sum / act.length).toFixed(1) : '0']
  ].map(([k, v]) => `<div class="kpi"><small>${k}</small><div>${v}</div></div>`).join('');

  $('ph').innerHTML = `<tr><th>프로그램</th><th class="n">상한(1인)</th><th class="n">참여 횟수</th><th class="n">참여 인원</th>` +
    S.gn.map(g => `<th class="n">${esc(g)}</th>`).join('') +
    `<th class="n">인정 포인트 합계</th><th class="n">상한 도달</th></tr>`;
  const mx = Math.max(1, ...S.pr.map(r => r.pts));
  $('pb').innerHTML = S.pr.map(r =>
    `<tr><td><b>${esc(r.name)}</b></td><td class="n">${esc(r.cap)}</td><td class="n">${r.cnt}</td><td class="n">${r.ppl}</td>` +
    r.g.map(v => `<td class="n">${v}</td>`).join('') +
    `<td class="n"><div class="bw"><span>${nf(r.pts)}</span><div class="bar"><i style="width:${r.pts / mx * 100}%"></i></div></div></td>` +
    `<td class="n">${esc(r.hit)}</td></tr>`).join('');

  const opt = (id, arr) => {
    const el = $(id), first = el.firstElementChild.outerHTML;
    el.innerHTML = first + [...new Set(arr.filter(Boolean))].sort().map(v => `<option>${esc(v)}</option>`).join('');
  };
  opt('fg', P.map(p => p.grp));
  opt('fd', P.map(p => p.dept));
  renderRank();
}

function capHit(j, v) { const c = S.progs[j].cap; return c !== null && v >= c; }

function renderRank() {
  const q = $('q').value.trim().toLowerCase(), g = $('fg').value, d = $('fd').value;
  const rows = S.profs.filter(p =>
    (!g || p.grp === g) && (!d || p.dept === d) && (!q || (p.name + p.id + p.dept).toLowerCase().includes(q)));
  const mx = Math.max(1, ...S.profs.map(p => p.total));
  $('rc').textContent = '(' + rows.length + '명)';
  $('rb').innerHTML = rows.map(p =>
    `<tr><td class="c">${p.rank}</td><td>${esc(p.dept)}</td><td>${esc(p.grp)}</td><td>${esc(p.id)}</td><td><b>${esc(p.name)}</b></td>` +
    `<td class="n"><div class="bw"><span>${p.total}</span><div class="bar"><i style="width:${p.total / mx * 100}%"></i></div></div></td>` +
    `<td class="n">${p.cnt}</td>` +
    `<td>${p.pts.map((v, j) => v > 0 ? `<span class="chip${capHit(j, v) ? ' hit' : ''}">${esc(S.progs[j].name)} ${v}</span>` : '').join('')}</td></tr>`
  ).join('') || '<tr><td colspan="8" class="mute">조건에 맞는 교수가 없습니다.</td></tr>';
}

/* ---------- 교수 조회 ---------- */
function lookup() {
  const id = $('lid').value.trim(), p = S.profs.find(x => x.id === id), o = $('lo');
  if (!id) { o.textContent = '교번을 입력해 주세요.'; return; }
  if (!p) { o.textContent = '해당 교번의 교수를 찾을 수 없습니다.'; return; }
  const m = Math.max(1, ...p.pts);
  o.innerHTML =
    `<div><b>${esc(p.name)}</b> · ${esc(p.dept)} · ${esc(p.grp)}</div>` +
    `<div class="big">${p.total}<small> 점 (전체 ${S.profs.length}명 중 ${p.rank}위)</small></div>` +
    (p.over > 0 ? `<div class="badge">상한 초과로 ${p.over}점은 인정되지 않았습니다.</div>` : '') +
    `<div style="margin-top:10px">` +
    S.progs.map((pr, j) => {
      const v = p.pts[j], cap = pr.cap, w = cap !== null ? Math.min(100, v / cap * 100) : v / m * 100;
      return `<div class="row"><div>${esc(pr.name)}</div><div class="bar"><i style="width:${w}%"></i></div>` +
        `<div class="val">${v}${cap !== null ? ' / ' + cap : ''} ${capHit(j, v) ? '<span class="badge">상한</span>' : ''}</div></div>`;
    }).join('') + `</div>`;
}

/* ---------- 파일 업로드 ---------- */
function load(file) {
  $('msg').textContent = '';
  if (!file) return;
  const fr = new FileReader();
  fr.onload = e => {
    try {
      S = parse(XLSX.read(e.target.result, { type: 'array' }));
      $('app').hidden = false;
      render();
      $('lo').textContent = '교번을 입력하면 총 인정 포인트와 프로그램별 적립 현황이 표시됩니다.';
    } catch (err) {
      $('app').hidden = true;
      $('msg').textContent = err.message;
    }
  };
  fr.onerror = () => { $('msg').textContent = '파일을 읽지 못했습니다.'; };
  fr.readAsArrayBuffer(file);
}

const drop = $('drop'), inp = $('f');
drop.onclick = () => inp.click();
inp.onchange = () => load(inp.files[0]);
drop.ondragover = e => { e.preventDefault(); drop.classList.add('on'); };
drop.ondragleave = () => drop.classList.remove('on');
drop.ondrop = e => { e.preventDefault(); drop.classList.remove('on'); load(e.dataTransfer.files[0]); };

['q', 'fg', 'fd'].forEach(id => $(id).addEventListener('input', () => S && renderRank()));
$('lb').onclick = () => S && lookup();
$('lid').addEventListener('keydown', e => { if (e.key === 'Enter' && S) lookup(); });

$('t1').onclick = () => { $('v1').hidden = false; $('v2').hidden = true; $('t1').classList.add('on'); $('t2').classList.remove('on'); };
$('t2').onclick = () => { $('v1').hidden = true; $('v2').hidden = false; $('t2').classList.add('on'); $('t1').classList.remove('on'); };

    /* ===================== Supabase 클라우드 연동 유틸 ===================== */
    let _supabaseClient = null;

    function getSupabaseConfig() {
      const localUrl = localStorage.getItem('WBS_SUPABASE_URL');
      const localKey = localStorage.getItem('WBS_SUPABASE_ANON_KEY');
      if (localUrl && localKey) {
        return { url: localUrl, anonKey: localKey };
      }
      if (window.WBS_SUPABASE_CONFIG && window.WBS_SUPABASE_CONFIG.url && window.WBS_SUPABASE_CONFIG.anonKey) {
        return window.WBS_SUPABASE_CONFIG;
      }
      return { url: '', anonKey: '' };
    }

    function getSupabaseClient() {
      if (_supabaseClient) return _supabaseClient;
      const config = getSupabaseConfig();
      if (config.url && config.anonKey) {
        try {
          _supabaseClient = supabase.createClient(config.url, config.anonKey);
          return _supabaseClient;
        } catch (e) {
          console.error("Supabase client init error:", e);
        }
      }
      return null;
    }

    function arrayBufferToBase64(buffer) {
      let binary = '';
      const bytes = new Uint8Array(buffer);
      const len = bytes.byteLength;
      for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      return window.btoa(binary);
    }

    function base64ToArrayBuffer(base64) {
      const binary_string = window.atob(base64);
      const len = binary_string.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
      }
      return bytes.buffer;
    }

    const DEFAULT_WBS_DATA = [];
    const DEFAULT_WBS_TODAY = (() => { const d = new Date(); return `${d.getFullYear()}-${('0'+(d.getMonth()+1)).slice(-2)}-${('0'+d.getDate()).slice(-2)}`; })();
    let WBS_DATA = DEFAULT_WBS_DATA;
    let CURRENT_SOURCE_NAME = "";

    // Legacy Storage implementation retained for backwards compatibility; new saves do not call it.
    function supabaseAddOrUpdateLegacy(record, existingId = null) {
      const client = getSupabaseClient();
      if (!client) return Promise.reject(new Error("Supabase not connected"));

      const bucketName = 'wbs-files';
      const filePath = `${record.wbsDate}/${record.fileName}`;
      const storageUri = `storage://${bucketName}/${filePath}`;

      let uploadPromise = Promise.resolve(storageUri);

      if (record.fileData) {
        const blob = new Blob([record.fileData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        uploadPromise = client.storage
          .from(bucketName)
          .upload(filePath, blob, {
            cacheControl: '3600',
            upsert: true
          })
          .then(({ data, error }) => {
            if (error) {
              console.error("Supabase Storage upload error:", error);
              throw error;
            }
            return storageUri;
          });
      } else {
        uploadPromise = Promise.resolve(null);
      }

      return uploadPromise.then(finalPath => {
        const dbRecord = {
          wbs_date: record.wbsDate,
          file_name: record.fileName,
          file_data: finalPath,
          tasks: record.tasks,
          summary: record.summary,
          saved_at: record.savedAt
        };

        if (existingId) {
          return client.from('wbs_analyses').update(dbRecord).eq('id', existingId).select('id')
            .then(({ data, error }) => {
              if (error) throw error;
              if (data && data.length > 0) return data[0].id;
              return client.from('wbs_analyses').insert(dbRecord).select('id').then(({ data: insData, error: insErr }) => {
                if (insErr) throw insErr;
                return insData[0].id;
              });
            });
        } else {
          return client.from('wbs_analyses').insert(dbRecord).select('id')
            .then(({ data, error }) => {
              if (error) throw error;
              return data[0].id;
            });
        }
      });
    }

    function supabaseAddOrUpdate(record, existingId = null) {
      const client = getSupabaseClient();
      if (!client) return Promise.reject(new Error('Supabase not connected'));

      // Small Excel workbooks are stored with the analysis record itself, avoiding
      // a separate Storage upload and making the saved result self-contained.
      const dbRecord = {
        wbs_date: record.wbsDate,
        file_name: record.fileName,
        file_data: record.fileData ? arrayBufferToBase64(record.fileData) : null,
        tasks: record.tasks,
        summary: record.summary,
        saved_at: record.savedAt
      };

      const save = existingId
        ? client.from('wbs_analyses').update(dbRecord).eq('id', existingId).select('id')
        : client.from('wbs_analyses').insert(dbRecord).select('id');

      return save.then(({ data, error }) => {
        if (error) throw error;
        if (data && data.length > 0) return data[0].id;

        // The selected record may have been removed by another user. Create it again.
        return client.from('wbs_analyses').insert(dbRecord).select('id').then(({ data: inserted, error: insertError }) => {
          if (insertError) throw insertError;
          return inserted[0].id;
        });
      });
    }

    function getCombinedHistoryList() {
      const client = getSupabaseClient();
      if (!client) {
        return dbGetAll();
      }
      
      return client.from('wbs_analyses').select('id, saved_at, wbs_date, file_name, summary').order('id', { ascending: false })
        .then(({ data, error }) => {
          if (error) {
            console.error("Supabase select error:", error);
            return dbGetAll();
          }
          
          const cloudList = (data || []).map(row => ({
            id: row.id,
            savedAt: row.saved_at,
            wbsDate: row.wbs_date,
            fileName: row.file_name,
            summary: row.summary
          }));
          
          return dbGetAll().then(localList => {
            const localMap = new Map(localList.map(r => [r.id, r]));
            cloudList.forEach(r => {
              const localRec = localMap.get(r.id);
              if (localRec && localRec.fileData) {
                r.fileData = localRec.fileData;
                r.tasks = localRec.tasks;
              }
            });
            return cloudList;
          });
        });
    }


    /* ===================== 날짜 유틸 ===================== */
    function parseDate(s) { if (!s) return null; const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
    function fmt(d) { if (!d) return '-'; const y = d.getFullYear(), m = ('0' + (d.getMonth() + 1)).slice(-2), dd = ('0' + d.getDate()).slice(-2); return `${y}-${m}-${dd}`; }
    function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
    function startOfDay(d) { const r = new Date(d); r.setHours(0, 0, 0, 0); return r; }

    function getWeekRange(date) {
      const d = startOfDay(date);
      const day = d.getDay();
      const diffToMon = (day === 0 ? -6 : 1 - day);
      const mon = addDays(d, diffToMon);
      const sun = addDays(mon, 6);
      return [mon, sun];
    }

    let TODAY = startOfDay(parseDate(DEFAULT_WBS_TODAY));

    function rangesOverlap(s1, e1, s2, e2) { return s1 <= e2 && s2 <= e1; }

    function isDone(t) {
      // 완료 여부는 실적 시트의 공정율 100%만 기준으로 삼습니다.
      // 계획 시트의 진척율/K열/N열, 실적 시트의 상태/완료일은 공정율 100%를 대체하지 않습니다.
      return t.hasActual && t.actualProgress >= 100;
    }

    function delayDays(t) {
      const end = parseDate(t.end);
      if (!end || isDone(t)) return 0;
      const diff = Math.round((TODAY - end) / 86400000);
      return diff > 0 ? diff : 0;
    }

    function isDelayByDate(t) {
      const end = parseDate(t.end);
      return !!(end && !isDone(t) && TODAY >= end);
    }

    function taskStatus(t) {
      if (isDone(t)) return 'done';
      if (isDelayByDate(t)) return 'delay';
      const start = parseDate(t.start), end = parseDate(t.end);
      if (start && start <= TODAY && (!end || end >= TODAY)) return 'progress';
      return 'upcoming';
    }
    const STATUS_LABEL = { done: '완료', delay: '지연', progress: '진행중', upcoming: '예정' };

    /* ===================== 표시 유틸 ===================== */
    const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

    // 엑셀·DB에서 온 문자열을 innerHTML에 넣기 전에 이스케이프합니다.
    function esc(v) {
      return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }
    // 화면 표기용 날짜(YYYY.MM.DD). 저장·비교용 값은 fmt()의 YYYY-MM-DD를 유지합니다.
    function fmtDot(d) { return d ? fmt(d).replace(/-/g, '.') : '-'; }
    function dotStr(s) { return s ? String(s).replace(/-/g, '.') : '-'; }
    function fmtRange(s, e) { return `${fmtDot(s)} ~ ${fmtDot(e)}`; }

    function emptyRow(colspan, msg) { return `<tr><td colspan="${colspan}" class="table-empty">${msg}</td></tr>`; }

    /* ===================== 렌더링 ===================== */
    const STATUS_BADGE = { done: 'badge-success', delay: 'badge-danger', progress: 'badge-warning', upcoming: 'badge-tag' };
    function badgeHtml(st) { return `<span class="${STATUS_BADGE[st]}">${STATUS_LABEL[st]}</span>`; }

    function taskPeriod(t) {
      const start = parseDate(t.start), end = parseDate(t.end);
      return start && end ? fmtRange(start, end) : '-';
    }

    const pctText = v => (v === undefined || v === null ? '-' : v);

    function diffHtml(t) {
      if (t.actualProgress === undefined || t.actualProgress === null) return '-';
      const diff = t.actualProgress - t.progress;
      if (diff === 0) return '0';
      return `<span class="${diff < 0 ? 'text-danger' : 'text-success'}">${diff > 0 ? '+' : ''}${diff}</span>`;
    }

    // 일정 지연(마감 도래·초과)을 먼저, 그다음 실적 미달을 표시합니다.
    function noteHtml(t) {
      const dd = delayDays(t);
      if (dd > 0) return `<span class="text-danger">${dd}일 지연</span>`;
      if (isDelayByDate(t)) return '<span class="text-danger">마감일 도래</span>';
      if (hasProgressGap(t)) return `<span class="text-warning">실적 ${t.progress - t.actualProgress}%p 미달</span>`;
      return '-';
    }

    function groupByPhase(list) {
      const map = new Map();
      list.forEach(t => {
        const key = t.phase || '기타';
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(t);
      });
      return map;
    }

    // 단계는 WBS 등장 순서, 같은 단계 안에서는 시작일 순으로 정렬합니다.
    function sortByPhaseThenStart(list) {
      const order = new Map();
      list.forEach(t => { const key = t.phase || '기타'; if (!order.has(key)) order.set(key, order.size); });
      return [...list].sort((a, b) =>
        (order.get(a.phase || '기타') - order.get(b.phase || '기타')) || (a.start || '').localeCompare(b.start || ''));
    }

    function taskTableRowHtml(t, i) {
      return `<tr>
        <td>${i + 1}</td>
        <td>${badgeHtml(taskStatus(t))}</td>
        <td class="cell-wrap is-narrow">${esc(t.phase || '-')}</td>
        <td class="cell-wrap is-narrow">${esc(t.activity || '-')}</td>
        <td class="cell-wrap">${esc(t.task)}</td>
        <td class="cell-wrap is-narrow">${esc(t.deliverable || '-')}</td>
        <td>${taskPeriod(t)}</td>
        <td>${t.progress}</td>
        <td>${pctText(t.actualProgress)}</td>
        <td>${diffHtml(t)}</td>
        <td>${noteHtml(t)}</td>
      </tr>`;
    }

    function renderTaskTable(tbodyId, list, emptyMsg) {
      document.getElementById(tbodyId).innerHTML = list.length
        ? sortByPhaseThenStart(list).map(taskTableRowHtml).join('')
        : emptyRow(11, emptyMsg);
    }

    function previewRowHtml(t) {
      return `<tr>
        <td>${badgeHtml(taskStatus(t))}</td>
        <td class="cell-wrap is-narrow">${esc(t.phase || '-')}</td>
        <td class="cell-wrap">${esc(t.task)}</td>
        <td>${taskPeriod(t)}</td>
        <td>${t.progress}</td>
        <td>${pctText(t.actualProgress)}</td>
        <td>${noteHtml(t)}</td>
      </tr>`;
    }

    // 대시보드 목록 위젯은 5건만 보여 주고 전체는 '더보기 +'로 이동합니다.
    const DASHBOARD_PREVIEW_ROWS = 5;

    function renderDashboardPreview(tbodyId, list, emptyMsg, sortFn) {
      const sorted = sortFn ? [...list].sort(sortFn) : [...list].sort((a, b) => (a.start || '').localeCompare(b.start || ''));
      document.getElementById(tbodyId).innerHTML = sorted.length
        ? sorted.slice(0, DASHBOARD_PREVIEW_ROWS).map(previewRowHtml).join('')
        : emptyRow(7, emptyMsg);
    }

    function getTasksInRange(rangeStart, rangeEnd, excludeUpcoming = false) {
      return WBS_DATA.filter(t => {
        if (isDone(t)) return false;
        const st = taskStatus(t);
        if (excludeUpcoming && st === 'upcoming') return false;
        const s = parseDate(t.start), e = parseDate(t.end);
        if (!s || !e) return false;
        return rangesOverlap(s, e, rangeStart, rangeEnd);
      });
    }

    function getThisWeekWorkTasks(rangeStart, rangeEnd) {
      return WBS_DATA.filter(t => {
        const start = parseDate(t.start), end = parseDate(t.end);
        if (!start || !end) return false;

        // 완료 업무는 실제 완료일(없으면 계획 완료일)을 기준으로 금주 실적으로 포함합니다.
        if (isDone(t)) {
          const completedOn = parseDate(t.actual_end) || end;
          return completedOn >= rangeStart && completedOn <= rangeEnd;
        }

        // 미완료 업무는 금주 일정과 겹치면서 이미 시작된 진행·지연 업무만 포함합니다.
        return taskStatus(t) !== 'upcoming' && rangesOverlap(start, end, rangeStart, rangeEnd);
      });
    }

    function hasProgressGap(t) {
      return t.actualProgress !== undefined && t.actualProgress < t.progress;
    }

    function getDelayTasks() {
      return WBS_DATA.filter(t => {
        if (isDone(t)) return false;
        if (isDelayByDate(t)) return true;                                      // 계획완료일 도래/초과
        if (taskStatus(t) === 'progress' && hasProgressGap(t)) return true;     // 실적 미달
        return false;
      }).sort((a, b) => {
        const ddA = delayDays(a), ddB = delayDays(b);
        if (ddA !== ddB) return ddB - ddA;                                       // 일정 지연 먼저
        const gapA = (a.progress - (a.actualProgress ?? a.progress));
        const gapB = (b.progress - (b.actualProgress ?? b.progress));
        return gapB - gapA;                                                      // 갭 큰 순
      });
    }

    /* ===================== WBS 전체 일정 트리 ===================== */
    const CHEVRON_SVG = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>';

    function statusChipsHtml(tasks) {
      const count = st => tasks.filter(t => taskStatus(t) === st).length;
      const done = count('done'), delay = count('delay'), prog = count('progress');
      return `<span class="badge-tag">${tasks.length}건</span>`
        + (done ? `<span class="badge-success">완료 ${done}</span>` : '')
        + (delay ? `<span class="badge-danger">지연 ${delay}</span>` : '')
        + (prog ? `<span class="badge-warning">진행 ${prog}</span>` : '');
    }

    // 막대 너비만 데이터로 정해지므로 인라인 width를 허용합니다(가이드 E15).
    function progressBarHtml(pct, fillClass = '') {
      return `<span class="stats-bar-cell"><span class="stats-bar" aria-hidden="true"><span class="stats-bar-fill ${fillClass}" style="width:${pct}%"></span></span><span class="tree-pct">${pct}%</span></span>`;
    }

    function avgPlanProgress(tasks) {
      return tasks.length ? Math.round(tasks.reduce((s, t) => s + t.progress, 0) / tasks.length) : 0;
    }

    function treeTaskRowHtml(t) {
      const st = taskStatus(t);
      const actualEnd = parseDate(t.actual_end);
      return `<tr>
        <td>${badgeHtml(st)}</td>
        <td class="cell-wrap">${esc(t.task)}</td>
        <td class="cell-wrap is-narrow">${esc(t.deliverable || '-')}</td>
        <td>${taskPeriod(t)}</td>
        <td>${st === 'done' && actualEnd ? fmtDot(actualEnd) : '-'}</td>
        <td>${t.progress}</td>
        <td>${pctText(t.actualProgress)}</td>
        <td>${noteHtml(t)}</td>
      </tr>`;
    }

    function renderWbsTree() {
      const phaseFilter = document.getElementById('filterPhase').value;
      const statusFilter = document.getElementById('filterStatus').value;
      const kw = document.getElementById('filterKeyword').value.trim().toLowerCase();
      const tree = document.getElementById('wbsTree');

      const rows = WBS_DATA.filter(t => {
        if (phaseFilter && t.phase !== phaseFilter) return false;
        if (statusFilter && taskStatus(t) !== statusFilter) return false;
        if (kw && !(t.task || '').toLowerCase().includes(kw)) return false;
        return true;
      });

      document.getElementById('wbsTreeSummary').textContent = `총 ${rows.length}건`;

      if (rows.length === 0) {
        const msg = WBS_DATA.length
          ? '조건에 맞는 TASK가 없습니다. 검색 조건을 바꾸거나 초기화해 주세요.'
          : '분석할 WBS 데이터가 없습니다. WBS 업데이트 메뉴에서 파일을 올려 주세요.';
        tree.innerHTML = `<div class="table-scroll-wrapper"><table class="data-table"><tbody>${emptyRow(1, msg)}</tbody></table></div>`;
        return;
      }

      // 단계 → 활동 → task 그룹핑 (원래 순서 유지)
      const phaseMap = new Map();
      rows.forEach(t => {
        const ph = t.phase || '기타';
        const ac = t.activity || '기타';
        if (!phaseMap.has(ph)) phaseMap.set(ph, new Map());
        const actMap = phaseMap.get(ph);
        if (!actMap.has(ac)) actMap.set(ac, []);
        actMap.get(ac).push(t);
      });

      // 검색 조건이 있으면 활동까지 펼쳐 결과를 바로 보여 줍니다. 단계는 항상 펼친 상태로 시작합니다.
      const hasFilter = !!(phaseFilter || statusFilter || kw);
      let html = '';

      phaseMap.forEach((actMap, phase) => {
        const phaseTasks = [...actMap.values()].flat();
        html += `
<div class="tree-phase open">
  <button type="button" class="tree-phase-hd" aria-expanded="true" onclick="toggleTreeNode(this)">
    <span class="tree-chevron">${CHEVRON_SVG}</span>
    <span class="tree-phase-name">${esc(phase)}</span>
    <span class="tree-meta">${statusChipsHtml(phaseTasks)}${progressBarHtml(avgPlanProgress(phaseTasks))}</span>
  </button>
  <div class="tree-phase-body">`;

        actMap.forEach((tasks, activity) => {
          html += `
    <div class="tree-activity${hasFilter ? ' open' : ''}">
      <button type="button" class="tree-act-hd" aria-expanded="${hasFilter}" onclick="toggleTreeNode(this)">
        <span class="tree-chevron">${CHEVRON_SVG}</span>
        <span class="tree-act-name">${esc(activity)}</span>
        <span class="tree-meta">${statusChipsHtml(tasks)}${progressBarHtml(avgPlanProgress(tasks))}</span>
      </button>
      <div class="tree-act-body">
        <div class="table-scroll-wrapper">
          <table class="data-table" aria-label="${esc(activity)} TASK 목록">
            <thead><tr><th>상태</th><th>TASK</th><th>산출물</th><th>기간</th><th>완료일</th><th>계획(%)</th><th>실적(%)</th><th>비고</th></tr></thead>
            <tbody>${tasks.map(treeTaskRowHtml).join('')}</tbody>
          </table>
        </div>
      </div>
    </div>`;
        });

        html += `</div></div>`;
      });

      tree.innerHTML = html;
    }

    function toggleTreeNode(button) {
      const open = button.parentElement.classList.toggle('open');
      button.setAttribute('aria-expanded', String(open));
    }

    function setTreeAll(open) {
      document.querySelectorAll('#wbsTree .tree-phase, #wbsTree .tree-activity').forEach(el => {
        el.classList.toggle('open', open);
        el.firstElementChild.setAttribute('aria-expanded', String(open));
      });
    }

    function resetWbsFilters() {
      document.getElementById('filterPhase').value = '';
      document.getElementById('filterStatus').value = '';
      document.getElementById('filterKeyword').value = '';
      renderAllTable();
    }

    function renderAllTable() { renderWbsTree(); }

    /* 계획 공정율: 오늘 기준으로 시작했어야 할 TASK 비중 (경과일 기준 선형 계획 진도) */
    function calcPlanRate() {
      const totalW = WBS_DATA.reduce((s, t) => s + (t.weight || 0), 0) || 1;
      let acc = 0;
      WBS_DATA.forEach(t => {
        const s = parseDate(t.start), e = parseDate(t.end);
        const w = t.weight || 0;
        if (!s || !e) return;
        if (TODAY >= e) { acc += w; }
        else if (TODAY >= s) {
          const span = (e - s) / 86400000 + 1;
          const passed = (TODAY - s) / 86400000 + 1;
          acc += w * Math.min(1, passed / span);
        }
      });
      return acc / totalW * 100;
    }
    /* 실적 공정율: 가중치 * 실적 진척율 (실적 시트 없으면 계획 진척율 사용) */
    function calcActualRate() {
      const totalW = WBS_DATA.reduce((s, t) => s + (t.weight || 0), 0) || 1;
      const acc = WBS_DATA.reduce((s, t) => {
        const prog = (t.actualProgress !== undefined ? t.actualProgress : t.progress) / 100;
        return s + (t.weight || 0) * prog;
      }, 0);
      return acc / totalW * 100;
    }

    function weightedPlanProgress(tasks) {
      const totalW = tasks.reduce((s, t) => s + (t.weight || 0), 0) || 1;
      return Math.round(tasks.reduce((s, t) => s + (t.weight || 0) * t.progress, 0) / totalW);
    }

    function renderPhaseSummary() {
      const count = (tasks, st) => tasks.filter(t => taskStatus(t) === st).length;
      const rowCells = tasks => {
        const delay = count(tasks, 'delay');
        return `<td>${tasks.length}</td>
          <td>${count(tasks, 'done')}</td>
          <td>${count(tasks, 'progress')}</td>
          <td>${delay ? `<span class="text-danger">${delay}</span>` : 0}</td>
          <td>${count(tasks, 'upcoming')}</td>
          <td>${progressBarHtml(weightedPlanProgress(tasks), delay ? 'fill-red' : '')}</td>`;
      };

      let html = '';
      groupByPhase(WBS_DATA).forEach((tasks, phase) => {
        html += `<tr><td class="cell-wrap is-narrow">${esc(phase)}</td>${rowCells(tasks)}</tr>`;
      });
      document.getElementById('phaseSummary').innerHTML = html
        || emptyRow(7, '분석할 WBS 데이터가 없습니다. WBS 업데이트 메뉴에서 파일을 올려 주세요.');
      document.getElementById('phaseSummaryFoot').innerHTML = WBS_DATA.length
        ? `<tr><td>합계</td>${rowCells(WBS_DATA)}</tr>`
        : '';
    }

    function renderAll() {
      const [wMon, wSun] = getWeekRange(TODAY);
      const [nMon] = getWeekRange(addDays(TODAY, 7));
      const nFri = addDays(nMon, 4);

      document.getElementById('todayLabel').textContent = `${fmtDot(TODAY)}(${WEEKDAYS[TODAY.getDay()]})`;
      const srcTag = document.getElementById('srcTag');
      srcTag.textContent = '데이터 출처: ' + (CURRENT_SOURCE_NAME || '-');
      srcTag.title = CURRENT_SOURCE_NAME || '';

      const projEnd = parseDate('2027-01-29');
      const dday = Math.round((projEnd - TODAY) / 86400000);
      const ddayEl = document.getElementById('ddayLabel');
      ddayEl.textContent = dday >= 0 ? `사업종료 D-${dday}` : `사업종료 D+${-dday}`;
      ddayEl.className = 'badge-dday' + (dday < 0 ? ' closed' : '');

      const thisLabel = fmtRange(wMon, wSun);
      const nextLabel = fmtRange(nMon, nFri);
      ['thisWeekRangeLabel', 'thisWeekRangeLabel2'].forEach(id => document.getElementById(id).textContent = thisLabel);
      ['nextWeekRangeLabel', 'nextWeekRangeLabel2'].forEach(id => document.getElementById(id).textContent = nextLabel);

      const thisWeekTasks = getThisWeekWorkTasks(wMon, wSun);     // 금주: 완료·진행·지연 업무 전체
      const nextWeekTasks = getTasksInRange(nMon, nFri);          // 차주: 월~금 예정 업무
      const delayTasks = getDelayTasks();

      renderTaskTable('thisWeekList', thisWeekTasks, '이번 주에 진행한 TASK가 없습니다.');
      renderTaskTable('nextWeekList', nextWeekTasks, '다음 주에 예정된 TASK가 없습니다.');
      renderTaskTable('delayList', delayTasks, '지연된 TASK가 없습니다. 계획대로 진행 중입니다.');

      renderDashboardPreview('dashThisPreview', thisWeekTasks, '이번 주에 진행한 TASK가 없습니다.');
      renderDashboardPreview('dashNextPreview', nextWeekTasks, '다음 주에 예정된 TASK가 없습니다.');
      renderDashboardPreview('dashDelayPreview', delayTasks, '지연된 TASK가 없습니다. 계획대로 진행 중입니다.', (a, b) => delayDays(b) - delayDays(a));

      // 사이드바·목록 건수
      document.getElementById('navCountThis').textContent = thisWeekTasks.length;
      document.getElementById('navCountNext').textContent = nextWeekTasks.length;
      document.getElementById('navCountDelay').textContent = delayTasks.length;
      document.getElementById('thisWeekCountLabel').textContent = `총 ${thisWeekTasks.length}건`;
      document.getElementById('nextWeekCountLabel').textContent = `총 ${nextWeekTasks.length}건`;
      document.getElementById('dashDelayMeta').textContent = `총 ${delayTasks.length}건`;
      const dateCnt = delayTasks.filter(t => isDelayByDate(t)).length;
      const gapCnt  = delayTasks.length - dateCnt;
      const delayBreakdown = dateCnt > 0 && gapCnt > 0
        ? `마감 도래/초과 ${dateCnt}건 · 실적 미달 ${gapCnt}건`
        : gapCnt > 0 ? `실적 미달 ${gapCnt}건` : `마감 도래/초과 ${dateCnt}건`;
      document.getElementById('delayCountLabel').textContent = `총 ${delayTasks.length}건 (${delayBreakdown})`;

      // KPI
      document.getElementById('kpiTotal').textContent = WBS_DATA.length;
      document.getElementById('kpiThis').textContent = thisWeekTasks.length;
      document.getElementById('kpiNext').textContent = nextWeekTasks.length;
      document.getElementById('kpiDelay').textContent = delayTasks.length;
      document.getElementById('kpiDone').textContent = WBS_DATA.filter(t => taskStatus(t) === 'done').length;
      document.getElementById('kpiDelayCard').classList.toggle('is-alert', delayTasks.length > 0);

      // 공정율 게이지
      const planRate = calcPlanRate();
      const actualRate = calcActualRate();
      const gap = actualRate - planRate;
      const gapState = gap < -0.5 ? 'behind' : (gap > 0.5 ? 'ahead' : 'onTrack');
      document.getElementById('phPlanRate').textContent = planRate.toFixed(1) + '%';
      document.getElementById('phActualRate').textContent = actualRate.toFixed(1) + '%';
      const gapEl = document.getElementById('phGapRate');
      gapEl.textContent = (gap >= 0 ? '+' : '') + gap.toFixed(1) + '%p';
      gapEl.className = 'tally-count' + (gapState === 'behind' ? ' text-danger' : gapState === 'ahead' ? ' text-success' : '');
      document.getElementById('phDelayCnt').textContent = delayTasks.length + '건';
      const mainLabel = document.getElementById('phMainLabel');
      mainLabel.textContent = { behind: '계획 대비 지연 중', ahead: '계획 대비 선행 중', onTrack: '계획대로 진행 중' }[gapState];
      mainLabel.className = { behind: 'badge-danger', ahead: 'badge-success', onTrack: 'badge-info' }[gapState];
      document.getElementById('gaugeLabel').textContent = actualRate.toFixed(1) + '%';
      document.getElementById('gaugeSvg').setAttribute('aria-label',
        `실적 공정율 ${actualRate.toFixed(1)}%, 계획 공정율 ${planRate.toFixed(1)}%`);
      const circumference = 339.292;
      const offset = circumference - (circumference * Math.min(100, actualRate) / 100);
      document.getElementById('gaugeFill').style.strokeDashoffset = offset;

      renderAllTable();
      renderPhaseSummary();
    }

    /* ===================== 필터 초기화 ===================== */
    function initFilters() {
      const sel = document.getElementById('filterPhase');
      sel.innerHTML = '<option value="">전체 단계</option>';
      const phases = [...new Set(WBS_DATA.map(t => t.phase))];
      phases.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p; opt.textContent = p;
        sel.appendChild(opt);
      });
    }

    document.getElementById('filterPhase').addEventListener('change', renderAllTable);
    document.getElementById('filterStatus').addEventListener('change', renderAllTable);
    document.getElementById('filterKeyword').addEventListener('input', renderAllTable);

    function toggleSidebar() {
      const collapsed = document.body.classList.toggle('sidebar-collapsed');
      const button = document.getElementById('btnMenuToggle');
      button.setAttribute('aria-expanded', String(!collapsed));
      button.setAttribute('aria-label', collapsed ? '사이드바 메뉴 펼치기' : '사이드바 메뉴 접기');
    }

    /* ===================== 모달 공통 ===================== */
    // 닫기 3종: 헤더 ×·취소 버튼, 오버레이 클릭, Esc(가장 위 모달부터)
    function openModal(id) {
      const modal = document.getElementById(id);
      modal.classList.add('show');
      modal.setAttribute('aria-hidden', 'false');
      document.body.classList.add('modal-open');
    }

    function closeModal(id) {
      const modal = document.getElementById(id);
      modal.classList.remove('show');
      modal.setAttribute('aria-hidden', 'true');
      if (!document.querySelector('.modal-backdrop.show')) document.body.classList.remove('modal-open');
    }

    document.querySelectorAll('.modal-backdrop').forEach(modal => {
      modal.addEventListener('click', event => { if (event.target === modal) closeModal(modal.id); });
    });

    function setFieldError(input, message) {
      const error = document.getElementById(input.getAttribute('aria-describedby'));
      if (error) error.textContent = message || '';
      if (message) input.setAttribute('aria-invalid', 'true');
      else input.removeAttribute('aria-invalid');
    }

    /* ===== 관리자 인증 ===== */
    const PROTECTED_PAGES = new Set(['p-members', 'p-mail']);
    let _authed = false;
    let _pendingPage = '';

    function requireAuth(pageId) {
      if (!PROTECTED_PAGES.has(pageId) || _authed) { navigateTo(pageId); return; }
      _pendingPage = pageId;
      const inp = document.getElementById('authPwInput');
      inp.type = 'password';
      inp.value = '';
      setFieldError(inp, '');
      const toggle = document.querySelector('.btn-password-toggle');
      toggle.setAttribute('aria-pressed', 'false');
      toggle.setAttribute('aria-label', '비밀번호 표시');
      openModal('authOverlay');
      setTimeout(() => inp.focus(), 80);
    }

    function submitAuth() {
      const inp = document.getElementById('authPwInput');
      if (inp.value === '4343') {
        _authed = true;
        closeAuthModal();
        navigateTo(_pendingPage);
      } else {
        setFieldError(inp, '비밀번호가 올바르지 않습니다. 다시 입력해 주세요.');
        inp.value = '';
        inp.focus();
      }
    }

    function closeAuthModal() {
      closeModal('authOverlay');
    }

    function toggleAuthPassword() {
      const input = document.getElementById('authPwInput');
      const button = document.querySelector('.btn-password-toggle');
      const show = input.type === 'password';
      input.type = show ? 'text' : 'password';
      button.setAttribute('aria-pressed', String(show));
      button.setAttribute('aria-label', show ? '비밀번호 숨기기' : '비밀번호 표시');
      input.focus();
    }

    /* ===================== 화면 이동 · 화면 매뉴얼 ===================== */
    const PAGE_META = {
      'p-dash': { title: '대시보드', group: '업무 알림', manual: [
        '기준일의 전체 공정율(가중치 반영)과 금주·차주·지연 TASK 수를 한눈에 보여 줍니다.',
        '요약 카드나 목록의 더보기 +를 누르면 해당 화면으로 이동합니다.',
        '오른쪽 위 기준일 선택에서 저장된 다른 분석 결과를 불러올 수 있습니다.'] },
      'p-this': { title: '금주 한 일', group: '업무 알림', manual: [
        '기준일이 속한 주(월~일)에 완료했거나 진행·지연 중인 TASK를 보여 줍니다.',
        '완료 TASK는 실제 완료일(없으면 계획 완료일)이 이번 주에 있을 때 포함됩니다.',
        '차이(%p)는 실적 진척율에서 계획 진척율을 뺀 값입니다.'] },
      'p-next': { title: '차주 할 일', group: '업무 알림', manual: [
        '다음 주 월~금 일정과 겹치는 미완료 TASK를 보여 줍니다.',
        '주간보고 텍스트 복사를 누르면 금주 한 일·차주 할 일·지연 업무를 메일 본문 형식으로 복사합니다.'] },
      'p-delay': { title: '지연 업무', group: '업무 알림', manual: [
        '계획 완료일이 기준일과 같거나 지났는데 완료되지 않은 TASK와, 진행 중이지만 실적이 계획보다 낮은 TASK를 보여 줍니다.',
        '완료 여부는 WBS(실적) 시트의 공정율 100%만 기준으로 판단합니다.'] },
      'p-all': { title: 'WBS 전체 일정', group: '전체 일정', manual: [
        '단계 → 활동 → TASK 순서로 전체 일정을 펼쳐 볼 수 있습니다.',
        '단계·진행상태·TASK명으로 검색하면 조건에 맞는 활동이 펼쳐진 상태로 표시됩니다.',
        '막대는 TASK별 계획 진척율의 평균입니다.'] },
      'p-phase': { title: '단계별 현황', group: '전체 일정', manual: [
        '단계별 TASK 수와 상태별 건수, 가중치를 반영한 평균 계획 진척율을 보여 줍니다.',
        '지연 TASK가 있는 단계는 막대가 빨간색으로 표시됩니다.'] },
      'p-upload': { title: 'WBS 업데이트', group: '데이터 관리', manual: [
        '매주 최신 상세일정계획 WBS 엑셀을 올리면 모든 화면이 다시 계산됩니다.',
        '분석 결과를 확인한 뒤 최종 저장을 누르면 분석 결과 조회 메뉴에 기록됩니다. 같은 파일명은 덮어씁니다.',
        'Supabase를 연동하면 다른 PC에서도 같은 분석 결과를 조회할 수 있습니다.'] },
      'p-history': { title: '분석 결과 조회', group: '데이터 관리', manual: [
        '최종 저장한 분석 결과를 기준일별로 조회합니다.',
        '불러오기는 해당 분석 결과를 대시보드에 적용하고, 링크 복사는 바로 열 수 있는 주소를 복사합니다.',
        '다운로드는 저장 당시의 원본 WBS 엑셀 파일을 내려받습니다.'] },
      'p-members': { title: '구성원 관리', group: '메일 발송 관리', manual: [
        '메일의 받는 사람·참조 목록에 쓰이는 프로젝트 구성원을 관리합니다.',
        '구성원 추가를 눌러 등록하고, 표의 수정·삭제로 정보를 바꿀 수 있습니다.'] },
      'p-mail': { title: '메일 보내기', group: '메일 발송 관리', manual: [
        '받는 사람과 참조를 구성원 목록에서 고르고 제목과 내용을 입력한 뒤 발송합니다.',
        '내용 오른쪽에서 기준일을 고르면 해당 분석 결과로 금주 한 일·차주 할 일·지연 업무 본문이 채워집니다.',
        'EmailJS 설정이 없으면 기기의 기본 메일 앱으로 발송합니다.'] }
    };
    let CURRENT_PAGE = 'p-dash';

    function setNavGroupOpen(group, open) {
      group.classList.toggle('open', open);
      group.querySelector('.nav-link').setAttribute('aria-expanded', String(open));
    }

    function navigateTo(pageId) {
      const meta = PAGE_META[pageId];
      if (!meta) return;
      CURRENT_PAGE = pageId;
      document.querySelectorAll('.page').forEach(p => p.classList.toggle('on', p.id === pageId));

      // 현재 화면의 2Depth만 active, 그 1Depth만 펼칩니다.
      document.querySelectorAll('.submenu-item').forEach(item => {
        const active = item.dataset.page === pageId;
        item.classList.toggle('active', active);
        const link = item.querySelector('a');
        if (active) link.setAttribute('aria-current', 'page');
        else link.removeAttribute('aria-current');
      });
      document.querySelectorAll('.sidebar-nav .nav-item').forEach(group => {
        const hasActive = !!group.querySelector('.submenu-item.active');
        group.classList.toggle('active', hasActive);
        setNavGroupOpen(group, hasActive);
      });

      document.getElementById('pageTitle').textContent = meta.title;
      document.getElementById('bcGroup').textContent = meta.group;
      document.getElementById('bcCurrent').textContent = meta.title;
      document.title = `WBS 일정분석 - ${meta.title} | KEPCO ES`;
      closePageManual();
      window.scrollTo(0, 0);

      if (pageId === 'p-history') renderHistoryPage();
      if (pageId === 'p-members') renderMembersPage();
      if (pageId === 'p-mail') renderMailPage();
    }

    function openPageManual() {
      const meta = PAGE_META[CURRENT_PAGE];
      const panel = document.getElementById('pageManualPanel');
      panel.innerHTML = `<strong>${meta.title} 이용 방법</strong><ul>${meta.manual.map(m => `<li>${m}</li>`).join('')}</ul>`;
      panel.hidden = false;
      const button = document.getElementById('btnPageManual');
      button.setAttribute('aria-expanded', 'true');
      button.setAttribute('aria-label', `${meta.title} 화면 매뉴얼 닫기`);
      window.scrollTo(0, 0);
    }

    function closePageManual() {
      const panel = document.getElementById('pageManualPanel');
      if (panel.hidden) return;
      panel.hidden = true;
      const button = document.getElementById('btnPageManual');
      button.setAttribute('aria-expanded', 'false');
      button.setAttribute('aria-label', `${PAGE_META[CURRENT_PAGE].title} 화면 매뉴얼 열기`);
    }

    document.getElementById('btnPageManual').addEventListener('click', () => {
      if (document.getElementById('pageManualPanel').hidden) openPageManual();
      else closePageManual();
    });

    document.addEventListener('click', event => {
      if (event.target.closest('#pageManualPanel, #btnPageManual, .btn-sidebar-action')) return;
      closePageManual();
    });

    // 인쇄와 PDF 저장은 브라우저 출력 흐름을 사용하며, 출력 CSS가 main 콘텐츠만 남깁니다.
    function printCurrentPage(mode) {
      closePageManual();
      if (mode === 'pdf') {
        showToast('인쇄 창의 대상에서 "PDF로 저장"을 선택해 주세요.', 'info');
        setTimeout(() => window.print(), 300);
      } else {
        window.print();
      }
    }

    document.addEventListener('keydown', event => {
      if (event.key !== 'Escape') return;
      const openModals = document.querySelectorAll('.modal-backdrop.show');
      if (openModals.length) { closeModal(openModals[openModals.length - 1].id); return; }
      closePageManual();
    });

    document.querySelectorAll('.sidebar-nav .nav-link').forEach(link => {
      link.addEventListener('click', () => {
        const group = link.closest('.nav-item');
        if (document.body.classList.contains('sidebar-collapsed')) {
          toggleSidebar();
          setNavGroupOpen(group, true);
          return;
        }
        setNavGroupOpen(group, !group.classList.contains('open'));
      });
    });

    document.querySelectorAll('.submenu-item').forEach(item => {
      item.querySelector('a').addEventListener('click', event => {
        event.preventDefault();
        requireAuth(item.dataset.page);
      });
    });

    document.querySelectorAll('[data-go]').forEach(el => {
      el.addEventListener('click', event => {
        event.preventDefault();
        requireAuth(el.dataset.go);
      });
    });

    function populateFileSelect(activeId) {
      getCombinedHistoryList().then(all => {
        const sel = document.getElementById('fileSelect');
        if (!sel) return;
        const dates = [...new Set(all.map(r => r.wbsDate))].sort((a, b) => b.localeCompare(a));
        sel.innerHTML = '<option value="">기준일 선택</option>' +
          dates.map(d => {
            const rec = all.find(r => r.wbsDate === d);
            return `<option value="${rec.id}"${rec.id === activeId ? ' selected' : ''}>${dotStr(d)}</option>`;
          }).join('');
      });
    }

    function onFileSelectChange(idStr) {
      if (!idStr) return;
      restoreAnalysis(Number(idStr));
    }

    /* ===================== 메일 보내기 ===================== */
    const _mailTo = new Set();
    const _mailCc = new Set();
    let mailEditor = null;

    function ensureMailEditor() {
      if (mailEditor) return mailEditor;
      const editorEl = document.getElementById('mailBody');
      if (!editorEl || !window.Quill) return null;
      mailEditor = new Quill(editorEl, {
        theme: 'snow',
        placeholder: '메일 내용을 입력하세요',
        modules: {
          toolbar: [
            [{ header: [1, 2, 3, false] }],
            ['bold', 'italic', 'underline', 'strike'],
            [{ color: [] }, { background: [] }],
            [{ list: 'ordered' }, { list: 'bullet' }],
            [{ align: [] }],
            ['link', 'clean']
          ]
        }
      });
      return mailEditor;
    }

    function getMailBody() {
      const editor = ensureMailEditor();
      if (!editor) return { text: '', html: '' };
      return {
        text: editor.getText().trim(),
        html: editor.root.innerHTML
      };
    }

    function focusMailEditor() {
      const editor = ensureMailEditor();
      if (editor) editor.focus();
    }

    function populateMailBaselineSelect() {
      const select = document.getElementById('mailBaselineSelect');
      if (!select) return;
      const selected = select.value || 'current';
      // Supabase 포함 전체 이력에서 기준일 목록 구성
      getCombinedHistoryList().then(all => {
        // 업로드된 파일 1개당 1개의 기준일만 표시 (fileName 기준 최신 저장본)
        const byFile = new Map();
        [...all]
          .sort((a, b) => (a.savedAt || '').localeCompare(b.savedAt || ''))
          .forEach(r => byFile.set(r.fileName, r));
        const records = [...byFile.values()].sort((a, b) => b.wbsDate.localeCompare(a.wbsDate));
        select.innerHTML = '<option value="current">기준일 선택 — 현재 분석 (' + fmtDot(TODAY) + ')</option>' +
          records.map(r => `<option value="${r.id}">${dotStr(r.wbsDate)} · ${esc(r.fileName)}</option>`).join('');
        select.value = [...select.options].some(o => o.value === selected) ? selected : 'current';
      });
    }

    // IndexedDB 우선, 없으면 Supabase에서 tasks 포함 전체 레코드 fetch
    function fetchFullRecord(id) {
      return dbGet(id).then(local => {
        if (local && local.tasks) return local;
        const client = getSupabaseClient();
        if (!client) return local || null;
        return client.from('wbs_analyses').select('*').eq('id', id).maybeSingle()
          .then(({ data, error }) => {
            if (error || !data) return local || null;
            return {
              id: data.id,
              savedAt: data.saved_at,
              wbsDate: data.wbs_date,
              fileName: data.file_name,
              tasks: data.tasks,
              summary: data.summary
            };
          });
      });
    }

    const MAIL_HEADER_NOTICE =
      '안녕하세요. 이희성 이사입니다.\n' +
      '금주 한 일, 차주 할 일, 그리고 지연 업무에 대해서 안내드립니다.\n\n' +
      '해당 일정을 확인하신 후 파트(PM/기획, 개발, UI/UX) 차주 수요일까지 주간보고에 넣을 내용을 메일로 제출 부탁드립니다.\n\n' +
      '메일 주소 : 이희성  leehee43@16block.com\n\n' +
      '-'.repeat(95) + '\n' +
      '1. 금주 진행 업무 : 금주(월~금) 진행한 업무 내역(자유 기재)\n' +
      '2. 차주 진행 업무 : 차주(월~금)에 진행할 업무 내역(자유 기재)\n' +
      '3. 일정계획의 업무별 진도율 : 비율로 표시(예 : 시스템인프라 분석(80%))\n' +
      '4. 지연 업무 및 사유 : 일정계획 대비 달성하지 못한 업무 목록 나열 + 캐치업 방안\n' +
      '5. 요청 및 이슈사항 : 주관기관 담당자에게 보고해야 할 요청 및 이슈사항(자유 기재)\n' +
      '-'.repeat(95) + '\n' +
      '차주에 해야할 업무를 파트별로 일정을 보고 확인하셔야 하나,\n' +
      '확인 차원에서 본 메일은 매주 금요일에 발송드리겠습니다.\n\n' +
      '항상 감사드립니다.\n\n';

    function setMailBodyWithNotice(editor, bodyText, toastMsg) {
      editor.setContents({
        ops: [
          { insert: MAIL_HEADER_NOTICE, attributes: { bold: true } },
          { insert: bodyText }
        ]
      });
      showToast(toastMsg, 'ok');
    }

    function fillMailBodyFromBaseline(value) {
      const editor = ensureMailEditor();
      if (!editor) { showToast('본문 에디터를 불러오지 못했습니다. 네트워크 연결을 확인하세요.', 'err'); return; }
      if (value === 'current') {
        const text = buildNextWeekText(WBS_DATA, TODAY, CURRENT_SOURCE_NAME);
        if (!text) { showToast('현재 분석 결과에 복사할 업무가 없습니다.', ''); return; }
        setMailBodyWithNotice(editor, text, `기준일 ${fmtDot(TODAY)}의 업무 텍스트를 본문에 넣었습니다.`);
        return;
      }
      fetchFullRecord(Number(value)).then(record => {
        if (!record) { showToast('선택한 분석 결과를 찾을 수 없습니다.', 'err'); return; }
        const baseline = startOfDay(parseDate(record.wbsDate));
        const text = buildNextWeekText(record.tasks, baseline, record.fileName);
        if (!text) { showToast('선택한 기준일에 복사할 업무가 없습니다.', ''); return; }
        setMailBodyWithNotice(editor, text, `기준일 ${dotStr(record.wbsDate)}의 업무 텍스트를 본문에 넣었습니다.`);
      }).catch(err => showToast('분석 결과 불러오기 오류: ' + err.message, 'err'));
    }

    function renderMailPage() {
      ensureMailEditor();
      populateMailBaselineSelect();
      renderEjsStatus();
      const subjectEl = document.getElementById('mailSubject');
      if (subjectEl && !subjectEl.value) {
        subjectEl.value = '안녕하세요. 이희성 이사입니다. (차주 주간보고서 작성요청의 건)';
      }
      mdbGetAll().then(members => {
        _mailTo.clear();
        _mailCc.clear();
        const toWrap = document.getElementById('mailToChips');
        const ccWrap = document.getElementById('mailCcChips');
        if (!toWrap || !ccWrap) return;

        // 구성원 '이희성'의 이메일을 보내는 사람 주소로 표시
        const sender = members.find(m => m.name === '이희성');
        if (sender) {
          const fromName = document.getElementById('mailFromName');
          const fromAddr = document.getElementById('mailFromAddr');
          if (fromName) fromName.textContent = sender.name;
          if (fromAddr) fromAddr.textContent = sender.email;
        }

        if (members.length === 0) {
          const msg = `<span class="chip-empty">등록된 구성원이 없습니다. <a href="#p-members" onclick="event.preventDefault();navigateTo('p-members')">구성원 관리</a>에서 먼저 추가해 주세요.</span>`;
          toWrap.innerHTML = msg;
          ccWrap.innerHTML = msg;
          return;
        }

        const chipHtml = (m, area) =>
          `<button type="button" class="mail-chip" id="${area}-${m.id}" aria-pressed="false" onclick="toggleMailChip(${m.id},'${area}')">
            <svg class="mail-chip-check" aria-hidden="true" viewBox="0 0 24 24"><path d="M20 6 9 17l-5-5"/></svg>${esc(m.name)}<span class="mail-chip-sub">${esc(m.role || '')}</span>
          </button>`;
        toWrap.innerHTML = members.map(m => chipHtml(m, 'to')).join('');
        ccWrap.innerHTML = members.map(m => chipHtml(m, 'cc')).join('');
      });
    }

    function toggleMailChip(id, area) {
      const set = area === 'to' ? _mailTo : _mailCc;
      const el = document.getElementById(`${area}-${id}`);
      if (!el) return;
      if (set.has(id)) set.delete(id);
      else set.add(id);
      el.classList.toggle('selected', set.has(id));
      el.setAttribute('aria-pressed', String(set.has(id)));
    }

    /* --- EmailJS 설정 관리 --- */
    function getEjsConfig() {
      // emailjs.config.js는 GitHub Pages 배포용 공통 설정입니다.
      // 값이 모두 있으면 로컬 설정보다 우선하여 모든 방문자에게 동일하게 적용합니다.
      const deployed = window.WBS_EMAILJS_CONFIG || {};
      const local = {
        publicKey:  localStorage.getItem('ejs_pub') || '',
        serviceId:  localStorage.getItem('ejs_svc') || '',
        templateId: localStorage.getItem('ejs_tpl') || '',
        fromName:   localStorage.getItem('ejs_from_name') || '',
        replyTo:    localStorage.getItem('ejs_reply_to') || ''
      };
      const isDeployed = !!(deployed.publicKey && deployed.serviceId && deployed.templateId);
      const source = isDeployed ? deployed : local;
      return {
        publicKey: source.publicKey || '',
        serviceId: source.serviceId || '',
        templateId: source.templateId || '',
        fromName: source.fromName || '프로젝트 일정관리',
        replyTo: source.replyTo || '',
        source: isDeployed ? 'deployment' : 'browser'
      };
    }
    function isEjsConfigured() {
      const c = getEjsConfig();
      // EmailJS 발송에 필수인 값은 공개 키·서비스 ID·템플릿 ID뿐입니다.
      // 발신 표시명과 회신 주소는 템플릿에서 선택적으로 사용할 수 있으므로,
      // 이전 버전에서 저장한 3개 항목 설정도 계속 사용할 수 있게 둡니다.
      return !!(c.publicKey && c.serviceId && c.templateId);
    }
    function isValidEmail(value) {
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
    }
    function saveEjsConfig() {
      const pub = document.getElementById('ejsPublicKey').value.trim();
      const svc = document.getElementById('ejsServiceId').value.trim();
      const tpl = document.getElementById('ejsTemplateId').value.trim();
      const fromName = document.getElementById('ejsFromName').value.trim();
      const replyTo = document.getElementById('ejsReplyTo').value.trim();
      if (!pub || !svc || !tpl) { showToast('Public Key, Service ID, Template ID를 모두 입력하세요.', 'err'); return; }
      if (replyTo && !isValidEmail(replyTo)) { showToast('회신 이메일 형식을 확인하세요.', 'err'); document.getElementById('ejsReplyTo').focus(); return; }
      localStorage.setItem('ejs_pub', pub);
      localStorage.setItem('ejs_svc', svc);
      localStorage.setItem('ejs_tpl', tpl);
      localStorage.setItem('ejs_from_name', fromName);
      localStorage.setItem('ejs_reply_to', replyTo);
      renderEjsStatus();
      toggleEjsPanel(false);
      showToast('EmailJS 설정이 저장되었습니다.', 'ok');
    }
    function clearEjsConfig() {
      if (!confirm('EmailJS 설정을 초기화하시겠습니까?')) return;
      ['ejs_pub','ejs_svc','ejs_tpl','ejs_from_name','ejs_reply_to'].forEach(k => localStorage.removeItem(k));
      document.getElementById('ejsPublicKey').value  = '';
      document.getElementById('ejsServiceId').value  = '';
      document.getElementById('ejsTemplateId').value = '';
      document.getElementById('ejsFromName').value   = '';
      document.getElementById('ejsReplyTo').value    = '';
      renderEjsStatus();
      showToast('설정이 초기화되었습니다.', '');
    }
    function renderEjsStatus() {
      const ok = isEjsConfigured();
      const badge = document.getElementById('ejsStatusBadge');
      const note  = document.getElementById('mailSendModeNote');
      if (badge) {
        badge.className = ok ? 'badge-success' : 'badge-warning';
        const cfg = getEjsConfig();
        badge.textContent = ok
          ? (cfg.source === 'deployment' ? '배포 설정 완료 — EmailJS 직접 발송' : '브라우저 설정 완료 — EmailJS 직접 발송')
          : '미설정 — 메일 앱으로 대체 발송';
      }
      if (note) {
        note.textContent = ok
          ? '※ EmailJS를 통해 직접 발송됩니다.'
          : '※ EmailJS 미설정 — 기기의 기본 메일 앱으로 발송됩니다.';
      }
      const cfg = getEjsConfig();
      document.getElementById('ejsPublicKey').value  = cfg.publicKey;
      document.getElementById('ejsServiceId').value  = cfg.serviceId;
      document.getElementById('ejsTemplateId').value = cfg.templateId;
      document.getElementById('ejsFromName').value   = cfg.fromName;
      document.getElementById('ejsReplyTo').value    = cfg.replyTo;
      const fromName = document.getElementById('mailFromName');
      const fromAddr = document.getElementById('mailFromAddr');
      if (fromName) fromName.textContent = cfg.fromName;
      if (fromAddr) fromAddr.textContent = cfg.replyTo || '회신 이메일 미설정';
    }
    function toggleEjsPanel(force) {
      const body = document.getElementById('ejsPanelBody');
      const open = typeof force === 'boolean' ? force : body.hidden;
      body.hidden = !open;
      const button = document.getElementById('btnEjsToggle');
      button.setAttribute('aria-expanded', String(open));
      button.textContent = open ? '설정 닫기' : '설정 열기';
    }

    /* --- 메일 발송 --- */
    function sendMail() {
      if (_mailTo.size === 0) { showToast('받는 사람을 선택하세요.', 'err'); return; }
      const subject = document.getElementById('mailSubject').value.trim();
      const body = getMailBody();
      if (!subject) { showToast('제목을 입력하세요.', 'err'); document.getElementById('mailSubject').focus(); return; }
      if (!body.text) { showToast('내용을 입력하세요.', 'err'); focusMailEditor(); return; }

      mdbGetAll().then(members => {
        const toEmails = [..._mailTo].map(id => members.find(m => m.id === id)?.email).filter(Boolean);
        const ccEmails = [..._mailCc].map(id => members.find(m => m.id === id)?.email)
          .filter(email => email && !toEmails.includes(email));
        if (toEmails.length === 0) { showToast('유효한 받는 사람 이메일이 없습니다.', 'err'); return; }
        if (isEjsConfigured()) {
          sendViaEmailJS(toEmails, ccEmails, subject, body.text, body.html);
        } else {
          sendViaMailto(toEmails, ccEmails, subject, body.text);
        }
      });
    }

    function sendViaEmailJS(toEmails, ccEmails, subject, body, bodyHtml) {
      const cfg = getEjsConfig();
      if (!window.emailjs) {
        showToast('EmailJS 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인하세요.', 'err');
        return;
      }
      const btn = document.querySelector('.btn-mail-send');
      const origHTML = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> 발송 중…';

      // init/send가 동기 예외를 내도 발송 버튼이 비활성 상태로 남지 않게 Promise로 감쌉니다.
      Promise.resolve().then(() => {
        emailjs.init({ publicKey: cfg.publicKey });
        return emailjs.send(cfg.serviceId, cfg.templateId, {
          to_email:  toEmails.join(', '),
          cc_email:  ccEmails.join(', '),
          from_name: cfg.fromName,
          reply_to:  cfg.replyTo,
          subject:   subject,
          message:   body,
          message_html: bodyHtml,
          sent_at: new Date().toLocaleString('ko-KR'),
          page_url: window.location.href
        });
      }).then(() => {
        showToast('메일이 성공적으로 발송되었습니다.', 'ok');
        resetMailForm();
      }).catch(err => {
        showToast('발송 실패: ' + getEmailJsErrorMessage(err), 'err');
      }).finally(() => {
        btn.disabled = false;
        btn.innerHTML = origHTML;
      });
    }

    function getEmailJsErrorMessage(err) {
      const raw = String((err && (err.text || err.message)) || JSON.stringify(err) || '알 수 없는 오류');
      const msg = raw.toLowerCase();
      if (msg.includes('origin') || msg.includes('domain') || msg.includes('block')) {
        return '도메인이 허용되지 않았습니다. EmailJS에 현재 주소(localhost 또는 GitHub Pages)를 허용 도메인으로 등록하세요.';
      }
      if (msg.includes('public key') || msg.includes('service') || msg.includes('template') || msg.includes('not found')) {
        return 'EmailJS Public Key, Service ID, Template ID와 템플릿 상태를 확인하세요.';
      }
      return raw;
    }

    function sendViaMailto(toEmails, ccEmails, subject, body) {
      const params = [];
      if (ccEmails.length) params.push('cc=' + encodeURIComponent(ccEmails.join(',')));
      params.push('subject=' + encodeURIComponent(subject));
      params.push('body='    + encodeURIComponent(body));
      const a = document.createElement('a');
      a.href = `mailto:${toEmails.join(',')}?${params.join('&')}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast('메일 앱을 실행합니다.', 'ok');
    }

    function resetMailForm() {
      _mailTo.clear();
      _mailCc.clear();
      document.querySelectorAll('#mailToChips .mail-chip, #mailCcChips .mail-chip').forEach(el => {
        el.classList.remove('selected');
        el.setAttribute('aria-pressed', 'false');
      });
      document.getElementById('mailSubject').value = '';
      const editor = ensureMailEditor();
      if (editor) editor.setText('');
    }

    /* ===================== 엑셀 업로드 / 파싱 ===================== */
    function excelDateToStr(v) {
      if (v instanceof Date) {
        const y = v.getFullYear(), m = ('0' + (v.getMonth() + 1)).slice(-2), d = ('0' + v.getDate()).slice(-2);
        return `${y}-${m}-${d}`;
      }
      // Excel 시리얼 숫자 → 날짜 변환 (Excel epoch: 1899-12-30 기준, 1900 윤년 버그 포함)
      if (typeof v === 'number' && v > 1 && v < 100000) {
        const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
        const y = d.getUTCFullYear(), mm = ('0' + (d.getUTCMonth() + 1)).slice(-2), dd = ('0' + d.getUTCDate()).slice(-2);
        return `${y}-${mm}-${dd}`;
      }
      // 문자열 날짜 (예: "2026-06-18")
      if (typeof v === 'string') {
        const m = v.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
        if (m) return `${m[1]}-${('0'+m[2]).slice(-2)}-${('0'+m[3]).slice(-2)}`;
      }
      return null;
    }

    function parseWbsWorkbook(workbook) {
      const planName = workbook.SheetNames.find(n => n.includes('계획')) || workbook.SheetNames[0];
      const actName  = workbook.SheetNames.find(n => n.includes('실적'));

      const ws   = workbook.Sheets[planName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, cellDates: true });

      // 실적 시트 (없으면 null)
      const actRows = actName
        ? XLSX.utils.sheet_to_json(workbook.Sheets[actName], { header: 1, raw: true, defval: null, cellDates: true })
        : null;

      // row[5][3] = "Today" 레이블, row[5][4] = 기준일 날짜
      const todayCell = (rows[5] || [])[4];
      const todayStr = excelDateToStr(todayCell) || null;

      const data = [];
      let curPhase = null, curActivity = null;
      for (let r = 9; r < rows.length; r++) {
        const row = rows[r] || [];
        const phase       = row[1];   // B열: 단계
        const activity    = row[2];   // C열: 활동
        const task        = row[3];   // D열: TASK
        const weight      = row[4];   // E열: 가중치
        const deliverable = row[5];   // F열: 산출물
        const progress    = row[6];   // G열: 계획 진척율 (0~1)
        const start       = row[11];  // L열: 시작
        const end         = row[12];  // M열: 계획완료

        if (phase)    curPhase    = String(phase).replace(/\r\n/g, ' ').trim();
        if (activity) curActivity = String(activity).replace(/\r\n/g, ' ').trim();
        if (!task || typeof task !== 'string') continue;

        // 실적 시트에서 동일 행 진척율 읽기
        const actRow = actRows ? (actRows[r] || []) : [];
        const actProg = actRow[6];
        const actStatus = actRow[10];  // 실적 시트 K열: 진행상태
        const actEnd = actRow[13];     // 실적 시트 N열: 실제완료
        const hasActual = typeof actProg === 'number';
        const planProgress   = Math.round((typeof progress === 'number' ? progress : 0) * 100);
        const actualProgress = hasActual
          ? Math.round(actProg * 100)
          : planProgress;   // 실적 시트 없으면 계획과 동일

        data.push({
          phase: curPhase, activity: curActivity, task: task.trim(),
          weight: typeof weight === 'number' ? weight : 0,
          deliverable: deliverable ? String(deliverable).replace(/\r\n/g, ' ') : '',
          progress: planProgress,
          actualProgress,
          hasActual,
          status: typeof actStatus === 'number' ? actStatus : 0,
          start: excelDateToStr(start),
          end: excelDateToStr(end),
          actual_end: excelDateToStr(actEnd)
        });
      }
      return { tasks: data, todayStr };
    }

    function addUploadHistory(filename, count, ok) {
      const tbody = document.getElementById('uploadHistoryBody');
      if (tbody.querySelector('.table-empty')) tbody.innerHTML = '';
      const now = new Date();
      const tstr = `${fmtDot(now)} ${('0' + now.getHours()).slice(-2)}:${('0' + now.getMinutes()).slice(-2)}`;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${tstr}</td><td class="cell-wrap">${esc(filename)}</td><td>${count}</td><td>${ok ? '<span class="badge-success">적용 완료</span>' : '<span class="badge-danger">실패</span>'}</td>`;
      tbody.prepend(tr);
    }

    function setUploadStatus(type, text) {
      const el = document.getElementById('uploadStatus');
      el.hidden = false;
      el.className = 'upload-status is-' + type;
      el.innerHTML = `<svg aria-hidden="true" viewBox="0 0 24 24">${STATUS_ICONS[type]}</svg><span></span>`;
      el.querySelector('span').textContent = text;
    }

    function handleFile(file) {
      setUploadStatus('info', `"${file.name}" 분석 중…`);

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const buffer = e.target.result;
          const data = new Uint8Array(buffer);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const { tasks: parsed, todayStr } = parseWbsWorkbook(workbook);
          if (parsed.length === 0) {
            throw new Error('인식된 TASK가 없습니다. "WBS(계획)" 시트 구조를 확인해 주세요.');
          }
          WBS_DATA = parsed;
          CURRENT_SOURCE_NAME = file.name;
          CURRENT_FILE_BUFFER = buffer;
          // 파일명 앞 8자리(YYYYMMDD)를 기준일로 우선 사용, 없으면 엑셀 Today 셀, 그것도 없으면 시스템 날짜
          const fileNameMatch = file.name.match(/^(\d{4})(\d{2})(\d{2})/);
          const fileDateStr = fileNameMatch ? `${fileNameMatch[1]}-${fileNameMatch[2]}-${fileNameMatch[3]}` : null;
          const resolvedDate = fileDateStr || todayStr;
          if (resolvedDate) {
            TODAY = startOfDay(parseDate(resolvedDate));
          }
          initFilters();
          renderAll();
          const dateNote = fileDateStr ? ` · 기준일 ${dotStr(fileDateStr)}(파일명)` : (todayStr ? ` · 기준일 ${dotStr(todayStr)}(엑셀 Today 셀)` : '');
          setUploadStatus('success', `"${file.name}" 적용 완료 — TASK ${parsed.length}건 인식${dateNote}`);
          showToast(`"${file.name}" 분석을 적용했습니다.`, 'ok');
          addUploadHistory(file.name, parsed.length, true);
          showSaveBar(file.name, fmt(TODAY), parsed.length);
        } catch (err) {
          setUploadStatus('danger', `파일을 분석하지 못했습니다: ${err.message}`);
          addUploadHistory(file.name, 0, false);
        }
      };
      reader.onerror = function () {
        setUploadStatus('danger', '파일을 읽는 중 오류가 발생했습니다.');
        addUploadHistory(file.name, 0, false);
      };
      reader.readAsArrayBuffer(file);
    }

    /* ===================== IndexedDB ===================== */
    const DB_NAME = 'WbsAnalysisDB', DB_VER = 2, STORE = 'analyses', MSTORE = 'members';
    let _db = null;

    function initDB() {
      return new Promise((resolve, reject) => {
        if (_db) { resolve(_db); return; }
        const req = indexedDB.open(DB_NAME, DB_VER);
        req.onupgradeneeded = e => {
          const d = e.target.result;
          if (!d.objectStoreNames.contains(STORE)) {
            const s = d.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
            s.createIndex('wbsDate', 'wbsDate', { unique: false });
          }
          if (!d.objectStoreNames.contains(MSTORE)) {
            d.createObjectStore(MSTORE, { keyPath: 'id', autoIncrement: true });
          }
        };
        req.onsuccess = e => { _db = e.target.result; resolve(_db); };
        req.onerror = e => reject(e.target.error);
      });
    }

    function dbGet(id) {
      return initDB().then(d => new Promise((resolve, reject) => {
        const req = d.transaction(STORE, 'readonly').objectStore(STORE).get(id);
        req.onsuccess = e => resolve(e.target.result);
        req.onerror = e => reject(e.target.error);
      }));
    }

    function dbGetAll() {
      return initDB().then(d => new Promise((resolve, reject) => {
        const req = d.transaction(STORE, 'readonly').objectStore(STORE).getAll();
        req.onsuccess = e => resolve([...e.target.result].reverse());
        req.onerror = e => reject(e.target.error);
      }));
    }

    function dbAdd(record) {
      return initDB().then(d => new Promise((resolve, reject) => {
        const req = d.transaction(STORE, 'readwrite').objectStore(STORE).add(record);
        req.onsuccess = e => resolve(e.target.result);
        req.onerror = e => reject(e.target.error);
      }));
    }

    function dbDelete(id) {
      return initDB().then(d => new Promise((resolve, reject) => {
        const req = d.transaction(STORE, 'readwrite').objectStore(STORE).delete(id);
        req.onsuccess = () => resolve();
        req.onerror = e => reject(e.target.error);
      }));
    }

    function dbUpdate(id, record) {
      return initDB().then(d => new Promise((resolve, reject) => {
        const rec = Object.assign({}, record, { id });
        const req = d.transaction(STORE, 'readwrite').objectStore(STORE).put(rec);
        req.onsuccess = () => resolve(id);
        req.onerror = e => reject(e.target.error);
      }));
    }

    /* === Members store (MSTORE) helpers === */
    function mdbGetAll() {
      const client = getSupabaseClient();
      if (!client) {
        return initDB().then(d => new Promise((resolve, reject) => {
          const req = d.transaction(MSTORE, 'readonly').objectStore(MSTORE).getAll();
          req.onsuccess = e => resolve(e.target.result);
          req.onerror = e => reject(e.target.error);
        }));
      }

      return client.from('wbs_members').select('*').order('name', { ascending: true })
        .then(({ data, error }) => {
          if (error) {
            console.error("Supabase members select error:", error);
            // Local fallback
            return initDB().then(d => new Promise((resolve, reject) => {
              const req = d.transaction(MSTORE, 'readonly').objectStore(MSTORE).getAll();
              req.onsuccess = e => resolve(e.target.result);
              req.onerror = e => reject(e.target.error);
            }));
          }

          const cloudList = (data || []).map(row => ({
            id: row.id,
            name: row.name,
            role: row.role,
            email: row.email,
            phone: row.phone,
            title: row.title,
            createdAt: row.created_at
          }));

          // Sync cache to IndexedDB
          return initDB().then(d => new Promise((resolve) => {
            const tx = d.transaction(MSTORE, 'readwrite');
            const store = tx.objectStore(MSTORE);
            store.clear().onsuccess = () => {
              if (cloudList.length === 0) {
                resolve(cloudList);
                return;
              }
              let count = 0;
              cloudList.forEach(m => {
                store.put(m).onsuccess = () => {
                  count++;
                  if (count === cloudList.length) resolve(cloudList);
                };
              });
            };
            tx.oncomplete = () => resolve(cloudList);
            tx.onerror = () => resolve(cloudList);
          }));
        });
    }
    function mdbAdd(record) {
      const client = getSupabaseClient();
      if (!client) {
        return initDB().then(d => new Promise((resolve, reject) => {
          const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).add(record);
          req.onsuccess = e => resolve(e.target.result);
          req.onerror = e => reject(e.target.error);
        }));
      }

      const dbRecord = {
        name: record.name,
        role: record.role,
        email: record.email,
        phone: record.phone,
        title: record.title,
        created_at: record.createdAt
      };

      return client.from('wbs_members').insert(dbRecord).select('id')
        .then(({ data, error }) => {
          if (error) throw error;
          const newId = data[0].id;
          const cachedRecord = Object.assign({}, record, { id: newId });
          // Save to IndexedDB cache
          return initDB().then(d => new Promise((resolve, reject) => {
            const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).put(cachedRecord);
            req.onsuccess = () => resolve(newId);
            req.onerror = e => reject(e.target.error);
          }));
        });
    }
    function mdbPut(id, record) {
      const client = getSupabaseClient();
      if (!client) {
        return initDB().then(d => new Promise((resolve, reject) => {
          const rec = Object.assign({}, record, { id });
          const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).put(rec);
          req.onsuccess = () => resolve(id);
          req.onerror = e => reject(e.target.error);
        }));
      }

      const dbRecord = {
        name: record.name,
        role: record.role,
        email: record.email,
        phone: record.phone,
        title: record.title,
        created_at: record.createdAt
      };

      return client.from('wbs_members').update(dbRecord).eq('id', id).select('id')
        .then(({ data, error }) => {
          if (error) throw error;
          const cachedRecord = Object.assign({}, record, { id });
          // Save to IndexedDB cache
          return initDB().then(d => new Promise((resolve, reject) => {
            const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).put(cachedRecord);
            req.onsuccess = () => resolve(id);
            req.onerror = e => reject(e.target.error);
          }));
        });
    }
    function mdbDelete(id) {
      const client = getSupabaseClient();
      if (!client) {
        return initDB().then(d => new Promise((resolve, reject) => {
          const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).delete(id);
          req.onsuccess = () => resolve();
          req.onerror = e => reject(e.target.error);
        }));
      }

      return client.from('wbs_members').delete().eq('id', id)
        .then(({ error }) => {
          if (error) throw error;
          return initDB().then(d => new Promise((resolve, reject) => {
            const req = d.transaction(MSTORE, 'readwrite').objectStore(MSTORE).delete(id);
            req.onsuccess = () => resolve();
            req.onerror = e => reject(e.target.error);
          }));
        });
    }

    function dbClear() {
      return initDB().then(d => new Promise((resolve, reject) => {
        const req = d.transaction(STORE, 'readwrite').objectStore(STORE).clear();
        req.onsuccess = () => resolve();
        req.onerror = e => reject(e.target.error);
      }));
    }

    function deleteAllAnalyses() {
      dbGetAll().then(all => {
        if (all.length === 0) { showToast('삭제할 저장 결과가 없습니다.', ''); return; }
        if (!confirm(`저장된 분석 결과 ${all.length}건을 모두 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.`)) return;
        dbClear().then(() => {
          showToast('전체 삭제 완료', 'ok');
          updateHistoryCount();
          populateFileSelect();
          renderHistoryPage();
          updateUrlParam(null);
        }).catch(err => showToast('삭제 오류: ' + err.message, 'err'));
      });
    }

    /* ===================== 토스트 ===================== */
    // 안내 박스·토스트·업로드 상태가 같은 4종 아이콘을 씁니다(가이드 5.9.1).
    const STATUS_ICONS = {
      success: '<circle cx="12" cy="12" r="9"/><path d="m8 12.5 3 3 5-6"/>',
      info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><path d="M12 8h.01"/>',
      warning: '<path d="M12 3 2.5 20h19L12 3Z"/><path d="M12 10v4"/><path d="M12 17h.01"/>',
      danger: '<circle cx="12" cy="12" r="9"/><path d="M12 8v5"/><path d="M12 16.5h.01"/>'
    };
    // 기존 호출부의 type('ok'·'err'·'info'·'')을 가이드의 토스트 종류로 바꿉니다.
    const TOAST_KIND = { ok: 'success', err: 'danger', warn: 'warning', info: 'info', '': 'info' };
    const TOAST_MAX = 4;

    // 성공·정보는 3초 뒤 사라지고, 오류·경고는 사용자가 닫을 때까지 유지합니다.
    function showToast(msg, type) {
      const kind = TOAST_KIND[type || ''] || 'info';
      const stack = document.getElementById('toastStack');
      const toast = document.createElement('div');
      toast.className = 'toast ' + kind;
      toast.setAttribute('role', kind === 'danger' || kind === 'warning' ? 'alert' : 'status');
      toast.innerHTML = `<svg class="toast-icon" aria-hidden="true" viewBox="0 0 24 24">${STATUS_ICONS[kind]}</svg><span class="toast-message"></span><button type="button" class="toast-close" aria-label="알림 닫기">&times;</button>`;
      toast.querySelector('.toast-message').textContent = msg;
      toast.querySelector('.toast-close').addEventListener('click', () => toast.remove());
      stack.appendChild(toast);
      while (stack.children.length > TOAST_MAX) stack.firstElementChild.remove();
      if (kind === 'success' || kind === 'info') setTimeout(() => toast.remove(), 3000);
    }

    /* ===================== 저장 관련 상태 ===================== */
    let CURRENT_FILE_BUFFER = null;

    function showSaveBar(fileName, wbsDate, taskCount) {
      const bar = document.getElementById('saveBar');
      document.getElementById('saveBarTitle').textContent = `"${fileName}" — TASK ${taskCount}건 인식, 기준일 ${dotStr(wbsDate)}`;
      document.getElementById('saveBarSub').textContent = '분석 결과를 확인한 후 최종 저장하면 분석 결과 조회 메뉴에서 다시 불러올 수 있습니다.';
      bar.classList.add('visible');
    }

    function saveCurrentAnalysis() {
      if (!CURRENT_FILE_BUFFER) {
        showToast('먼저 WBS 파일을 업로드하세요.', 'err'); return;
      }
      const planRate = calcPlanRate();
      const actualRate = calcActualRate();
      const record = {
        savedAt: new Date().toISOString(),
        wbsDate: fmt(TODAY),
        fileName: CURRENT_SOURCE_NAME,
        fileData: CURRENT_FILE_BUFFER.slice(0),
        tasks: JSON.parse(JSON.stringify(WBS_DATA)),
        summary: {
          total: WBS_DATA.length,
          done: WBS_DATA.filter(t => taskStatus(t) === 'done').length,
          delay: WBS_DATA.filter(t => taskStatus(t) === 'delay').length,
          inProgress: WBS_DATA.filter(t => taskStatus(t) === 'progress').length,
          upcoming: WBS_DATA.filter(t => taskStatus(t) === 'upcoming').length,
          planRate: +planRate.toFixed(1),
          actualRate: +actualRate.toFixed(1)
        }
      };

      const client = getSupabaseClient();
      if (client) {
        getCombinedHistoryList().then(all => {
          const existing = all.find(r => r.fileName === CURRENT_SOURCE_NAME);
          const existingId = existing ? existing.id : null;
          
          if (existingId && !confirm(`같은 파일이 이미 등록되어 있습니다.\n업데이트 됩니다. 계속 진행하시겠습니까?`)) return;
          
          showToast('클라우드에 저장하는 중...', 'info');
          supabaseAddOrUpdate(record, existingId).then(globalId => {
            dbUpdate(globalId, record).then(() => {
              showToast('클라우드 및 로컬 저장 완료!', 'ok');
              updateHistoryCount();
              populateFileSelect(globalId);
              renderHistoryPage();
              updateUrlParam(globalId);
            }).catch(err => showToast('로컬 캐시 저장 오류: ' + err.message, 'err'));
          }).catch(err => {
            console.error('클라우드 저장 오류:', err);
            // 클라우드 설정/권한 문제로 저장이 실패해도 사용자가 만든 분석 결과는
            // 브라우저에 남겨서 조회와 재시도가 가능하도록 한다.
            dbGetAll().then(localAll => {
              const localExisting = localAll.find(r => r.fileName === CURRENT_SOURCE_NAME);
              const saveLocal = localExisting ? dbUpdate(localExisting.id, record) : dbAdd(record);
              return saveLocal.then(id => {
                showToast('클라우드 저장에 실패하여 이 브라우저에 저장했습니다: ' + err.message, '');
                updateHistoryCount();
                populateFileSelect(id);
                renderHistoryPage();
                updateUrlParam(id);
              });
            }).catch(localErr => showToast('저장 오류: ' + localErr.message, 'err'));
          });
        });
      } else {
        dbGetAll().then(all => {
          const existing = all.find(r => r.fileName === CURRENT_SOURCE_NAME);
          if (existing) {
            if (!confirm(`같은 파일이 이미 등록되어 있습니다.\n업데이트 됩니다. 계속 진행하시겠습니까?`)) return;
            dbUpdate(existing.id, record).then(id => {
              showToast('로컬 저장 완료 (클라우드 미연동)', 'ok');
              updateHistoryCount();
              populateFileSelect(id);
              renderHistoryPage();
              updateUrlParam(id);
            }).catch(err => showToast('저장 오류: ' + err.message, 'err'));
          } else {
            dbAdd(record).then(id => {
              showToast('로컬 저장 완료 (클라우드 미연동)', 'ok');
              updateHistoryCount();
              populateFileSelect(id);
              updateUrlParam(id);
            }).catch(err => showToast('저장 오류: ' + err.message, 'err'));
          }
        });
      }
    }

    /* ===================== 분석 결과 조회 페이지 ===================== */
    function updateHistoryCount() {
      getCombinedHistoryList().then(list => {
        const badge = document.getElementById('navCountHistory');
        badge.textContent = list.length;
        badge.hidden = list.length === 0;
      });
    }

    /* === Supabase UI Event Handlers === */
    function updateSupabaseStatus() {
      const config = getSupabaseConfig();
      const badge = document.getElementById('supabaseStatusBadge');
      const urlInput = document.getElementById('sbUrl');
      const keyInput = document.getElementById('sbKey');
      
      if (urlInput) urlInput.value = config.url || '';
      if (keyInput) keyInput.value = config.anonKey || '';

      const client = getSupabaseClient();
      if (client) {
        badge.textContent = '연동됨';
        badge.className = 'badge-success';
      } else {
        badge.textContent = '연동 안 됨';
        badge.className = 'badge-tag';
      }
    }

    function initSupabaseEvents() {
      const btnSave = document.getElementById('btnSbSave');
      const btnClear = document.getElementById('btnSbClear');
      
      if (btnSave) {
        btnSave.addEventListener('click', () => {
          const url = document.getElementById('sbUrl').value.trim();
          const key = document.getElementById('sbKey').value.trim();
          
          if (!url || !key) {
            showToast('API URL과 Anon Key를 모두 입력하세요.', 'err');
            return;
          }
          
          localStorage.setItem('WBS_SUPABASE_URL', url);
          localStorage.setItem('WBS_SUPABASE_ANON_KEY', key);
          _supabaseClient = null;
          
          const testClient = getSupabaseClient();
          if (testClient) {
            testClient.from('wbs_analyses').select('id').limit(1).then(({ error }) => {
              if (error) {
                showToast('연동 실패: ' + error.message, 'err');
                localStorage.removeItem('WBS_SUPABASE_URL');
                localStorage.removeItem('WBS_SUPABASE_ANON_KEY');
                _supabaseClient = null;
                updateSupabaseStatus();
              } else {
                showToast('Supabase 클라우드 연동 성공!', 'ok');
                updateSupabaseStatus();
                updateHistoryCount();
                populateFileSelect();
                renderHistoryPage();
              }
            }).catch(err => {
              showToast('연동 실패: ' + err.message, 'err');
              localStorage.removeItem('WBS_SUPABASE_URL');
              localStorage.removeItem('WBS_SUPABASE_ANON_KEY');
              _supabaseClient = null;
              updateSupabaseStatus();
            });
          } else {
            showToast('Supabase 클라이언트 생성 실패.', 'err');
            updateSupabaseStatus();
          }
        });
      }
      
      if (btnClear) {
        btnClear.addEventListener('click', () => {
          if (!confirm('Supabase 클라우드 설정을 초기화하시겠습니까? (로컬 IndexedDB 데이터는 보존됩니다)')) return;
          localStorage.removeItem('WBS_SUPABASE_URL');
          localStorage.removeItem('WBS_SUPABASE_ANON_KEY');
          _supabaseClient = null;
          document.getElementById('sbUrl').value = '';
          document.getElementById('sbKey').value = '';
          showToast('클라우드 연동이 초기화되었습니다.', '');
          updateSupabaseStatus();
          updateHistoryCount();
          populateFileSelect();
          renderHistoryPage();
        });
      }
    }

    function resetHistSearch() {
      document.getElementById('histSearchName').value = '';
      renderHistoryPage(1);
    }

    /* ===================== 프로젝트 구성원 관리 ===================== */
    const MEMBER_FIELDS = [
      { id: 'mfName',  label: '이름',   key: 'name' },
      { id: 'mfTitle', label: '직책',   key: 'title' },
      { id: 'mfRole',  label: '역할',   key: 'role' },
      { id: 'mfPhone', label: '연락처', key: 'phone' },
      { id: 'mfEmail', label: '이메일', key: 'email' }
    ];

    function toggleMemberForm(editData) {
      document.getElementById('memberFormTitle').textContent = editData ? '구성원 수정' : '구성원 추가';
      document.getElementById('mfEditId').value = editData ? editData.id : '';
      document.getElementById('mfCreatedAt').value = editData ? (editData.createdAt || '') : '';
      MEMBER_FIELDS.forEach(f => {
        const el = document.getElementById(f.id);
        el.value = editData ? (editData[f.key] || '') : '';
        setFieldError(el, '');
      });
      openModal('memberModal');
      setTimeout(() => document.getElementById('mfName').focus(), 80);
    }

    function closeMemberForm() {
      closeModal('memberModal');
    }

    function saveMember() {
      let firstInvalid = null;
      MEMBER_FIELDS.forEach(f => {
        const el = document.getElementById(f.id);
        const value = el.value.trim();
        let message = value ? '' : `${f.label}을(를) 입력해 주세요.`;
        if (!message && f.key === 'email' && !isValidEmail(value)) message = '이메일 형식을 확인해 주세요.';
        setFieldError(el, message);
        if (message && !firstInvalid) firstInvalid = el;
      });
      if (firstInvalid) { firstInvalid.focus(); return; }

      const name  = document.getElementById('mfName').value.trim();
      const email = document.getElementById('mfEmail').value.trim();
      const editId = document.getElementById('mfEditId').value;
      const createdAt = document.getElementById('mfCreatedAt').value || new Date().toISOString();
      const rec = {
        name,
        role:  document.getElementById('mfRole').value.trim(),
        email,
        phone: document.getElementById('mfPhone').value.trim(),
        title: document.getElementById('mfTitle').value.trim(),
        createdAt
      };

      const promise = editId ? mdbPut(Number(editId), rec) : mdbAdd(rec);
      promise.then(() => {
        showToast(editId ? `"${name}" 정보가 수정되었습니다.` : `"${name}" 구성원이 추가되었습니다.`, 'ok');
        closeMemberForm();
        renderMembersPage();
      }).catch(err => showToast('저장 오류: ' + err.message, 'err'));
    }

    function editMember(id) {
      mdbGetAll().then(all => {
        const m = all.find(r => r.id === id);
        if (m) toggleMemberForm(m);
      });
    }

    function deleteMember(id) {
      mdbGetAll().then(all => {
        const m = all.find(r => r.id === id);
        if (!m) return;
        if (!confirm(`"${m.name}" (${m.email}) 구성원을 삭제하시겠습니까?`)) return;
        mdbDelete(id).then(() => {
          showToast('삭제되었습니다.', '');
          renderMembersPage();
        }).catch(err => showToast('삭제 오류: ' + err.message, 'err'));
      });
    }

    function renderMembersPage() {
      const tbody = document.getElementById('memberTableBody');
      const countLabel = document.getElementById('memberCountLabel');
      const navBadge = document.getElementById('navCountMembers');
      if (!tbody) return;

      mdbGetAll().then(all => {
        const count = all.length;
        if (countLabel) countLabel.textContent = count;
        if (navBadge) { navBadge.textContent = count; navBadge.hidden = count === 0; }

        if (count === 0) {
          tbody.innerHTML = emptyRow(7, '등록된 구성원이 없습니다. 구성원 추가 버튼을 눌러 등록해 주세요.');
          return;
        }

        tbody.innerHTML = all.map((m, i) => `<tr>
            <td>${i + 1}</td>
            <td>${esc(m.name)}</td>
            <td>${esc(m.role || '-')}</td>
            <td>${esc(m.title || '-')}</td>
            <td>${esc(m.email)}</td>
            <td>${esc(m.phone || '-')}</td>
            <td><div class="row-actions">
              <button type="button" class="btn-task-done" onclick="editMember(${m.id})">수정</button>
              <button type="button" class="btn-row-del" onclick="deleteMember(${m.id})">삭제</button>
            </div></td>
          </tr>`).join('');
      });
    }

    function fmtSavedAt(iso) {
      const d = new Date(iso);
      return `${fmtDot(d)} ${('0'+d.getHours()).slice(-2)}:${('0'+d.getMinutes()).slice(-2)}`;
    }

    /* 페이징: << < 1 … 10 > >> (10개 단위) */
    function renderPagination(navId, current, total, onPage) {
      const nav = document.getElementById(navId);
      const blockStart = Math.floor((current - 1) / 10) * 10 + 1;
      const blockEnd = Math.min(total, blockStart + 9);
      const navButton = (label, page, aria, disabled) =>
        `<button type="button" class="btn-page-nav" data-page="${page}" aria-label="${aria}"${disabled ? ' disabled' : ''}>${label}</button>`;
      let html = navButton('&laquo;', 1, '첫 페이지', current === 1) + navButton('&lsaquo;', current - 1, '이전 페이지', current === 1);
      for (let p = blockStart; p <= blockEnd; p++) {
        html += `<button type="button" class="btn-page-num${p === current ? ' active' : ''}" data-page="${p}" aria-label="${p}페이지"${p === current ? ' aria-current="page"' : ''}>${p}</button>`;
      }
      html += navButton('&rsaquo;', current + 1, '다음 페이지', current === total) + navButton('&raquo;', total, '마지막 페이지', current === total);
      nav.innerHTML = html;
      nav.onclick = event => {
        const button = event.target.closest('button[data-page]');
        if (button && !button.disabled) onPage(Number(button.dataset.page));
      };
    }

    // 분석 이력은 IndexedDB/클라우드에서 비동기로 불러온다. 이전 요청이 늦게
    // 끝나면 방금 고른 날짜의 결과를 다시 전체 목록으로 덮어쓰지 않도록 한다.
    let historyRenderVersion = 0;
    let historyPage = 1;

    function renderHistoryPage(page) {
      if (page) historyPage = page;
      const renderVersion = ++historyRenderVersion;
      const sel = document.getElementById('histSearchName');
      if (!sel) return;
      const selectedDate = sel.value;
      const tbody = document.getElementById('historyList');
      const summary = document.getElementById('historyCountLabel');

      getCombinedHistoryList().then(all => {
        if (renderVersion !== historyRenderVersion) return;
        const badge = document.getElementById('navCountHistory');
        if (badge) { badge.textContent = all.length; badge.hidden = all.length === 0; }

        // SELECT 옵션 갱신 (wbsDate 오름차순, 중복 제거)
        const dates = [...new Set(all.map(r => r.wbsDate))].sort((a, b) => a.localeCompare(b));
        sel.innerHTML = '<option value="">전체 날짜 보기</option>' +
          dates.map(d => `<option value="${d}"${d === selectedDate ? ' selected' : ''}>${dotStr(d)}</option>`).join('');

        const list = all.filter(r => !selectedDate || r.wbsDate === selectedDate)
          .sort((a, b) => a.wbsDate.localeCompare(b.wbsDate));
        const perPage = Number(document.getElementById('histPerPage').value) || 10;
        const totalPages = Math.max(1, Math.ceil(list.length / perPage));
        historyPage = Math.min(Math.max(1, historyPage), totalPages);
        summary.textContent = `${historyPage} / ${totalPages} (총 ${list.length}개)`;
        renderPagination('historyPagination', historyPage, totalPages, p => renderHistoryPage(p));

        if (list.length === 0) {
          tbody.innerHTML = emptyRow(11, all.length === 0
            ? '저장된 분석 결과가 없습니다. WBS 업데이트 후 최종 저장을 눌러 기록을 남겨 주세요.'
            : '조회된 데이터가 없습니다. 검색 조건을 바꾸거나 초기화해 주세요.');
          return;
        }

        const currentWbsDate = fmt(TODAY);
        const start = (historyPage - 1) * perPage;
        tbody.innerHTML = list.slice(start, start + perPage).map((r, idx) => {
          const isActive = r.wbsDate === currentWbsDate && r.fileName === CURRENT_SOURCE_NAME;
          const s = r.summary || {};
          const gap = (s.actualRate || 0) - (s.planRate || 0);
          const gapClass = gap < -0.5 ? 'text-danger' : (gap > 0.5 ? 'text-success' : '');
          const gapStr = (gap >= 0 ? '+' : '') + gap.toFixed(1);
          return `<tr${isActive ? ' class="is-selected" aria-current="true"' : ''}>
            <td>${start + idx + 1}</td>
            <td><span class="row-actions">${dotStr(r.wbsDate)}${isActive ? '<span class="badge-info">현재 분석 중</span>' : ''}</span></td>
            <td class="cell-wrap is-narrow" title="${esc(r.fileName)}">${esc(r.fileName)}</td>
            <td>${fmtSavedAt(r.savedAt)}</td>
            <td>${pctText(s.total)}</td>
            <td>${pctText(s.done)}</td>
            <td>${s.delay ? `<span class="text-danger">${s.delay}</span>` : pctText(s.delay)}</td>
            <td>${pctText(s.planRate)}</td>
            <td>${pctText(s.actualRate)}</td>
            <td>${gapClass ? `<span class="${gapClass}">${gapStr}</span>` : gapStr}</td>
            <td><div class="row-actions">
              <button type="button" class="btn-task-register" onclick="restoreAnalysis(${r.id})">불러오기</button>
              <button type="button" class="btn-task-done" onclick="copyAnalysisLink(${r.id})">링크 복사</button>
              <button type="button" class="btn-file-download" onclick="downloadAnalysisFile(${r.id})" aria-label="${esc(r.fileName)} 다운로드">다운로드</button>
              <button type="button" class="btn-row-del" onclick="deleteAnalysis(${r.id})">삭제</button>
            </div></td>
          </tr>`;
        }).join('');
      });
    }

    function updateUrlParam(id) {
      try {
        const url = new URL(window.location.href);
        if (id) url.searchParams.set('id', id);
        else url.searchParams.delete('id');
        history.replaceState(null, '', url.toString());
      } catch(e) {}
    }

    function buildNextWeekText(data, baseline, sourceName) {
      const taskDelayDays = t => {
        const end = parseDate(t.end);
        if (!end || isDone(t)) return 0;
        return Math.max(0, Math.round((baseline - end) / 86400000));
      };
      const isDelayByBaseline = t => {
        const end = parseDate(t.end);
        return !!(end && !isDone(t) && baseline >= end);
      };
      const isInProgress = t => {
        const start = parseDate(t.start), end = parseDate(t.end);
        return !isDone(t) && start && start <= baseline && (!end || end >= baseline);
      };
      const taskStatusFor = t => {
        if (isDone(t)) return 'done';
        if (isDelayByBaseline(t)) return 'delay';
        return isInProgress(t) ? 'progress' : 'upcoming';
      };

      // 1. 금주 한 일
      const [wMon, wSun] = getWeekRange(baseline);
      const thisWeekTasks = data.filter(t => {
        const start = parseDate(t.start), end = parseDate(t.end);
        if (!start || !end) return false;
        if (isDone(t)) {
          const completedOn = parseDate(t.actual_end) || end;
          return completedOn >= wMon && completedOn <= wSun;
        }
        return taskStatusFor(t) !== 'upcoming' && rangesOverlap(start, end, wMon, wSun);
      });

      // 2. 차주 할 일
      const [nMon, nSun] = getWeekRange(addDays(baseline, 7));
      const nFri = addDays(nMon, 4);
      const nextWeekTasks = data.filter(t => {
        if (isDone(t)) return false;
        const start = parseDate(t.start), end = parseDate(t.end);
        return start && end && rangesOverlap(start, end, nMon, nFri);
      });

      // 3. 지연 업무
      const delayTasks = data.filter(t => {
        if (isDone(t)) return false;
        return isDelayByBaseline(t) || (isInProgress(t) && hasProgressGap(t));
      }).sort((a, b) => taskDelayDays(b) - taskDelayDays(a) || ((b.progress - (b.actualProgress ?? b.progress)) - (a.progress - (a.actualProgress ?? a.progress))));

      if (thisWeekTasks.length === 0 && nextWeekTasks.length === 0 && delayTasks.length === 0) return '';

      const lines = [];

      // [금주 한 일] 렌더링
      const wRangeStr = `${fmt(wMon)} ~ ${fmt(wSun)}`;
      lines.push(`[금주 한 일] ${wRangeStr}`);
      lines.push('');
      if (thisWeekTasks.length === 0) {
        lines.push('이번 주에 진행한 업무가 없습니다.');
      } else {
        let lastPhase = '', lastAct = '';
        thisWeekTasks.forEach((t, i) => {
          if (t.phase && t.phase !== lastPhase) {
            lines.push(`▣ ${t.phase}`);
            lastPhase = t.phase; lastAct = '';
          }
          if (t.activity && t.activity !== lastAct) {
            lines.push(`  ◆ ${t.activity}`);
            lastAct = t.activity;
          }
          const st = STATUS_LABEL[taskStatusFor(t)] || '';
          const pct = t.progress != null ? ` (계획 ${t.progress}%)` : '';
          const ap = t.actualProgress != null ? ` / 실적 ${t.actualProgress}%` : '';
          const start = t.start ? ` ${t.start}` : '';
          const end = t.end ? ` ~ ${t.end}` : '';
          lines.push(`    ${i + 1}. [${st}] ${t.task}${pct}${ap} |${start}${end}`);
        });
      }

      lines.push('');

      // [차주 할 일] 렌더링
      const nRangeStr = `${fmt(nMon)} ~ ${fmt(nFri)}`;
      lines.push(`[차주 할 일] ${nRangeStr}`);
      lines.push('');
      if (nextWeekTasks.length === 0) {
        lines.push('다음 주에 예정된 TASK가 없습니다.');
      } else {
        let lastPhase = '', lastAct = '';
        nextWeekTasks.forEach((t, i) => {
          if (t.phase && t.phase !== lastPhase) {
            lines.push(`▣ ${t.phase}`);
            lastPhase = t.phase; lastAct = '';
          }
          if (t.activity && t.activity !== lastAct) {
            lines.push(`  ◆ ${t.activity}`);
            lastAct = t.activity;
          }
          const st = STATUS_LABEL[taskStatusFor(t)] || '';
          const pct = t.progress != null ? ` (계획 ${t.progress}%)` : '';
          const ap  = t.actualProgress != null && t.actualProgress !== t.progress
            ? ` / 실적 ${t.actualProgress}%` : '';
          const start = t.start ? ` ${t.start}` : '';
          const end   = t.end   ? ` ~ ${t.end}` : '';
          lines.push(`    ${i + 1}. [${st}] ${t.task}${pct}${ap} |${start}${end}`);
        });
      }

      lines.push('');

      // [지연 업무] 렌더링
      lines.push(`[지연 업무] 기준일 ${fmt(baseline)}`);
      lines.push('');
      if (delayTasks.length === 0) {
        lines.push('지연 업무가 없습니다.');
      } else {
        let lastDelayPhase = '', lastDelayAct = '';
        delayTasks.forEach((t, i) => {
          if (t.phase && t.phase !== lastDelayPhase) {
            lines.push(`▣ ${t.phase}`);
            lastDelayPhase = t.phase; lastDelayAct = '';
          }
          if (t.activity && t.activity !== lastDelayAct) {
            lines.push(`  ◆ ${t.activity}`);
            lastDelayAct = t.activity;
          }
          const dd = taskDelayDays(t);
          const gap = t.actualProgress != null ? t.progress - t.actualProgress : 0;
          const issue = isDelayByBaseline(t)
            ? (dd > 0 ? `${dd}일 지연` : '마감일 도래')
            : `실적 ${gap}%p 미달`;
          const pct = t.progress != null ? ` (계획 ${t.progress}%)` : '';
          const ap = t.actualProgress != null ? ` / 실적 ${t.actualProgress}%` : '';
          const start = t.start ? ` ${t.start}` : '';
          const end = t.end ? ` ~ ${t.end}` : '';
          lines.push(`    ${i + 1}. [${issue}] ${t.task}${pct}${ap} |${start}${end}`);
        });
      }

      lines.push('');
      lines.push(`※ 기준일: ${fmt(baseline)} / 출처: ${sourceName}`);
      return lines.join('\n');
    }

    function copyNextWeekText() {
      const text = buildNextWeekText(WBS_DATA, TODAY, CURRENT_SOURCE_NAME);
      if (!text) { showToast('복사할 주간 업무보고 내용이 없습니다.', ''); return; }
      const [wMon, wSun] = getWeekRange(TODAY);
      const [nMon] = getWeekRange(addDays(TODAY, 7));
      const nFri = addDays(nMon, 4);
      const thisTasks = getThisWeekWorkTasks(wMon, wSun);
      const nextTasks = getTasksInRange(nMon, nFri);
      const delayTasks = getDelayTasks();

      navigator.clipboard.writeText(text)
        .then(() => showToast(`금주 ${thisTasks.length}건 · 차주 ${nextTasks.length}건 · 지연 ${delayTasks.length}건 복사 완료`, 'ok'))
        .catch(() => { prompt('아래 내용을 복사하세요:', text); });
    }

    function copyAnalysisLink(id) {
      try {
        const url = new URL(window.location.href);
        url.searchParams.set('id', id);
        navigator.clipboard.writeText(url.toString())
          .then(() => showToast('링크가 클립보드에 복사되었습니다.', 'ok'))
          .catch(() => {
            prompt('아래 URL을 복사하세요:', url.toString());
          });
      } catch(e) { showToast('URL 생성 오류', 'err'); }
    }

    function restoreAnalysis(id) {
      dbGet(id).then(r => {
        if (r) {
          loadAnalysisRecord(r, id);
        } else {
          const client = getSupabaseClient();
          if (client) {
            showToast('클라우드에서 데이터를 불러오는 중...', 'info');
            client.from('wbs_analyses').select('*').eq('id', id).maybeSingle().then(({ data, error }) => {
              if (error) throw error;
              if (!data) {
                showToast('저장 레코드를 찾을 수 없습니다.', 'err');
                return;
              }
              
              let fileDataPromise = Promise.resolve(null);
              if (data.file_data) {
                if (data.file_data.startsWith('storage://')) {
                  const parts = data.file_data.replace('storage://', '').split('/');
                  const bucket = parts[0];
                  const path = parts.slice(1).join('/');
                  fileDataPromise = client.storage.from(bucket).download(path)
                    .then(({ data: blob, error: dlErr }) => {
                      if (dlErr) throw dlErr;
                      return blob.arrayBuffer();
                    });
                } else {
                  fileDataPromise = Promise.resolve(base64ToArrayBuffer(data.file_data));
                }
              }

              fileDataPromise.then(arrayBuffer => {
                const record = {
                  id: data.id,
                  savedAt: data.saved_at,
                  wbsDate: data.wbs_date,
                  fileName: data.file_name,
                  fileData: arrayBuffer,
                  tasks: data.tasks,
                  summary: data.summary
                };
                dbUpdate(data.id, record).then(() => {
                  loadAnalysisRecord(record, data.id);
                  updateHistoryCount();
                  populateFileSelect(data.id);
                }).catch(() => {
                  loadAnalysisRecord(record, data.id);
                });
              }).catch(err => {
                console.error("Storage file download error:", err);
                const record = {
                  id: data.id,
                  savedAt: data.saved_at,
                  wbsDate: data.wbs_date,
                  fileName: data.file_name,
                  fileData: null,
                  tasks: data.tasks,
                  summary: data.summary
                };
                loadAnalysisRecord(record, data.id);
              });
            }).catch(err => showToast('클라우드 불러오기 오류: ' + err.message, 'err'));
          } else {
            showToast('저장 레코드를 찾을 수 없습니다 (클라우드 미연동).', 'err');
          }
        }
      }).catch(err => showToast('불러오기 오류: ' + err.message, 'err'));
    }

    function loadAnalysisRecord(r, id) {
      let tasks = r.tasks;
      // 예전 버전에서 저장된 분석 결과는 계획 시트의 완료값을 들고 있을 수 있습니다.
      // 원본 엑셀 데이터가 있으면 다시 파싱해 완료 여부를 실적 시트 기준으로 보정합니다.
      if (r.fileData && window.XLSX) {
        try {
          const workbook = XLSX.read(new Uint8Array(r.fileData), { type: 'array', cellDates: true });
          const reparsed = parseWbsWorkbook(workbook);
          if (reparsed.tasks && reparsed.tasks.length > 0) tasks = reparsed.tasks;
        } catch (e) {
          console.warn('Saved workbook reparse failed:', e);
        }
      }

      WBS_DATA = tasks;
      CURRENT_SOURCE_NAME = r.fileName;
      CURRENT_FILE_BUFFER = r.fileData;
      TODAY = startOfDay(parseDate(r.wbsDate));
      initFilters();
      renderAll();
      showSaveBar(r.fileName, r.wbsDate, tasks.length);
      showToast(`"${r.fileName}" (기준일 ${dotStr(r.wbsDate)}) 불러오기 완료`, 'ok');
      updateUrlParam(id);
      populateFileSelect(id);
      navigateTo('p-dash');
    }

    function downloadAnalysisFile(id) {
      dbGet(id).then(r => {
        if (r && r.fileData) {
          triggerFileDownload(r.fileName, r.fileData);
        } else {
          const client = getSupabaseClient();
          if (client) {
            showToast('클라우드에서 파일 데이터를 가져오는 중...', 'info');
            client.from('wbs_analyses').select('file_name, file_data').eq('id', id).maybeSingle().then(({ data, error }) => {
              if (error) throw error;
              if (!data || !data.file_data) {
                showToast('저장된 파일 데이터가 없습니다.', 'err');
                return;
              }

              let fileDataPromise = Promise.resolve(null);
              if (data.file_data.startsWith('storage://')) {
                const parts = data.file_data.replace('storage://', '').split('/');
                const bucket = parts[0];
                const path = parts.slice(1).join('/');
                fileDataPromise = client.storage.from(bucket).download(path)
                  .then(({ data: blob, error: dlErr }) => {
                    if (dlErr) throw dlErr;
                    return blob.arrayBuffer();
                  });
              } else {
                fileDataPromise = Promise.resolve(base64ToArrayBuffer(data.file_data));
              }

              fileDataPromise.then(buffer => {
                triggerFileDownload(data.file_name, buffer);
              }).catch(err => showToast('파일 다운로드 실패: ' + err.message, 'err'));
            }).catch(err => showToast('다운로드 오류: ' + err.message, 'err'));
          } else {
            showToast('저장된 파일 데이터가 없습니다.', 'err');
          }
        }
      }).catch(err => showToast('다운로드 오류: ' + err.message, 'err'));
    }

    function triggerFileDownload(fileName, fileBuffer) {
      const blob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fileName;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast(`"${fileName}" 다운로드 시작`, 'ok');
    }

    function deleteAnalysis(id) {
      dbGet(id).then(r => {
        const labelName = r ? `"${r.fileName}" (기준일 ${r.wbsDate})` : `ID ${id}번 레코드`;
        if (!confirm(`${labelName} 저장 결과를 삭제하시겠습니까?`)) return;

        const client = getSupabaseClient();
        if (client) {
          showToast('클라우드에서 삭제 중...', 'info');
          client.from('wbs_analyses').delete().eq('id', id).then(({ error }) => {
            if (error) throw error;
            dbDelete(id).then(() => {
              showToast('삭제되었습니다.', '');
              updateHistoryCount();
              populateFileSelect();
              renderHistoryPage();
            }).catch(err => showToast('로컬 캐시 삭제 오류: ' + err.message, 'err'));
          }).catch(err => showToast('클라우드 삭제 오류: ' + err.message, 'err'));
        } else {
          dbDelete(id).then(() => {
            showToast('삭제되었습니다.', '');
            updateHistoryCount();
            populateFileSelect();
            renderHistoryPage();
          }).catch(err => showToast('삭제 오류: ' + err.message, 'err'));
        }
      });
    }

    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    uploadZone.addEventListener('click', () => fileInput.click());
    uploadZone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
    });
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
    });
    ['dragenter', 'dragover'].forEach(evt => {
      uploadZone.addEventListener(evt, (e) => { e.preventDefault(); uploadZone.classList.add('dragover'); });
    });
    ['dragleave', 'drop'].forEach(evt => {
      uploadZone.addEventListener(evt, (e) => { e.preventDefault(); uploadZone.classList.remove('dragover'); });
    });
    uploadZone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    });

    document.getElementById('btnSaveFinal').addEventListener('click', saveCurrentAnalysis);

    /* ===================== 초기 실행 ===================== */
    initFilters();
    renderAll();
    navigateTo('p-dash');
    initDB().then(() => {
      updateSupabaseStatus();
      initSupabaseEvents();
      updateHistoryCount();
      populateFileSelect();
      renderMembersPage();
      const initId = new URLSearchParams(window.location.search).get('id');
      if (initId) {
        restoreAnalysis(Number(initId));
      } else {
        // 최초 진입 시에는 저장된 WBS 중 Today(기준일)가 가장 최근인 파일을 기본으로 불러옵니다.
        // 같은 기준일이 여러 건이면 가장 나중에 저장한 파일을 선택합니다.
        getCombinedHistoryList().then(all => {
          const latest = all
            .filter(r => r.wbsDate && parseDate(r.wbsDate))
            .sort((a, b) => {
              const byWbsDate = b.wbsDate.localeCompare(a.wbsDate);
              return byWbsDate || String(b.savedAt || '').localeCompare(String(a.savedAt || ''));
            })[0];
          if (latest) restoreAnalysis(latest.id);
          else if (WBS_DATA.length === 0) navigateTo('p-upload');
        }).catch(() => {
          if (WBS_DATA.length === 0) navigateTo('p-upload');
        });
      }
    }).catch(() => {
      if (WBS_DATA.length === 0) navigateTo('p-upload');
    });

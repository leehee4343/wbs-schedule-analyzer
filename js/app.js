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
    function fmtShort(d) { if (!d) return '-'; const m = d.getMonth() + 1, dd = d.getDate(); const w = ['일', '월', '화', '수', '목', '금', '토'][d.getDay()]; return `${m}/${dd}(${w})`; }
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

    /* ===================== 렌더링 ===================== */
    function badgeHtml(st) { return `<span class="badge ${st}">${STATUS_LABEL[st]}</span>`; }

    function taskRowHtml(t, showPlanActual = false) {
      const st = taskStatus(t);
      const start = parseDate(t.start), end = parseDate(t.end);
      const dateLabel = (start && end) ? `${fmtShort(start)} ~ ${fmtShort(end)}` : '-';
      const barClass = st === 'delay' ? 'delaybar' : (st === 'done' ? 'donebar' : '');
      const dd = delayDays(t);
      const ap = t.actualProgress;
      const progGap = (ap !== undefined && ap < t.progress) ? t.progress - ap : 0;
      const delayNote = dd > 0
        ? `<span class="t-delaydays">${dd}일 지연</span>`
        : (st === 'delay' && isDelayByDate(t) ? `<span class="t-delaydays">마감일 도래</span>` : '');
      const progGapNote = (progGap > 0 && dd === 0)
        ? `<span class="t-delaydays" style="background:#FEF3C7;color:#92400E;border-color:#FDE68A;">실적 ${progGap}%p 미달</span>`
        : '';
      const phasePart = t.phase ? `<span class="m-tag">${t.phase}</span>` : '';
      const actPart = t.activity ? `<span style="color:var(--text-sub);">${t.activity}</span>` : '';
      const delivPart = t.deliverable ? `<span class="m-tag" style="background:var(--bg);color:var(--text-muted);">산출물: ${t.deliverable}</span>` : '';

      const hasActual = ap !== undefined && ap !== null;
      const hasDiff = hasActual && ap !== t.progress;
      let progressHtml;
      if (showPlanActual || hasDiff) {
        const actualValue = hasActual ? ap : 0;
        const diff = actualValue - t.progress;
        const gapClass = diff < 0 ? 'behind' : 'ahead';
        const gapStr = (diff > 0 ? '+' : '') + diff + '%p';
        const gapBadge = hasActual ? `<span class="gap-badge ${gapClass}">${gapStr}</span>` : '<span class="gap-badge behind">실적 미입력</span>';
        progressHtml = `
    <div class="dual-bar-wrap">
      <div class="dual-bar-row">
        <span class="dual-bar-label">계획</span>
        <div class="dual-bar-track"><div class="dual-bar-fill plan-fill" style="width:${t.progress}%"></div></div>
        <span class="dual-bar-pct">${t.progress}%</span>
      </div>
      <div class="dual-bar-sep"></div>
      <div class="dual-bar-row">
        <span class="dual-bar-label">실적</span>
        <div class="dual-bar-track"><div class="dual-bar-fill actual-fill" style="width:${actualValue}%"></div></div>
        <span class="dual-bar-pct">${hasActual ? actualValue + '%' : '-'}</span>
      </div>
      ${gapBadge}
    </div>`;
      } else {
        progressHtml = `
    <div class="task-bar-wrap">
      <div class="task-bar-bg ${barClass}"><div class="task-bar-fill" style="width:${t.progress}%;"></div></div>
      <div class="task-pct">${t.progress}%</div>
    </div>`;
      }

      return `
  <div class="task-row ${st === 'delay' ? 'is-delay' : ''}">
    ${badgeHtml(st)}
    <div class="task-info">
      <div class="t-name">${t.task}</div>
      <div class="t-meta">${phasePart}${actPart}${delivPart}${delayNote}${progGapNote}</div>
    </div>
    ${progressHtml}
    <div class="task-date">${dateLabel}</div>
  </div>`;
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

    function renderTaskGroup(container, list, emptyMsg, emptyGood, showPlanActual = false) {
      if (list.length === 0) {
        container.innerHTML = `<div class="empty-msg ${emptyGood ? 'good' : ''}">${emptyMsg}</div>`;
        return;
      }
      const grouped = groupByPhase(list);
      let html = '';
      grouped.forEach((tasks, phase) => {
        html += `<div class="phase-group"><div class="phase-title">${phase}</div>`;
        tasks.sort((a, b) => (a.start || '').localeCompare(b.start || ''));
        tasks.forEach(t => html += taskRowHtml(t, showPlanActual));
        html += `</div>`;
      });
      container.innerHTML = html;
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

    function renderDashboardPreview(container, list, max, emptyMsg, sortFn) {
      if (list.length === 0) {
        container.innerHTML = `<div class="empty-msg good">${emptyMsg}</div>`;
        return;
      }
      const sorted = sortFn ? [...list].sort(sortFn) : [...list].sort((a, b) => (a.start || '').localeCompare(b.start || ''));
      const top = sorted.slice(0, max);
      let html = '';
      top.forEach(t => html += taskRowHtml(t));
      if (sorted.length > max) {
        html += `<div class="empty-msg">외 ${sorted.length - max}건 더 있음</div>`;
      }
      container.innerHTML = html;
    }

    const CHEVRON_SVG = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg>`;

    function renderWbsTree() {
      const phaseFilter = document.getElementById('filterPhase').value;
      const statusFilter = document.getElementById('filterStatus').value;
      const kw = document.getElementById('filterKeyword').value.trim().toLowerCase();

      const rows = WBS_DATA.filter(t => {
        if (phaseFilter && t.phase !== phaseFilter) return false;
        const st = taskStatus(t);
        if (statusFilter && st !== statusFilter) return false;
        if (kw && !(t.task || '').toLowerCase().includes(kw)) return false;
        return true;
      });

      if (rows.length === 0) {
        document.getElementById('wbsTree').innerHTML = `<div class="empty-msg">조건에 맞는 TASK가 없습니다.</div>`;
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

      const hasFilter = !!(phaseFilter || statusFilter || kw);
      let html = '';

      phaseMap.forEach((actMap, phase) => {
        const phaseTasks = [...actMap.values()].flat();
        const total = phaseTasks.length;
        const done  = phaseTasks.filter(t => taskStatus(t) === 'done').length;
        const delay = phaseTasks.filter(t => taskStatus(t) === 'delay').length;
        const prog  = phaseTasks.filter(t => taskStatus(t) === 'progress').length;
        const avgProg = total ? Math.round(phaseTasks.reduce((s, t) => s + t.progress, 0) / total) : 0;

        const isOpen = hasFilter ? 'open' : 'open'; // 기본 단계 펼침
        html += `
<div class="tree-phase ${isOpen}" id="ph_${phase.replace(/\s/g,'_')}">
  <div class="tree-phase-hd" onclick="this.closest('.tree-phase').classList.toggle('open')">
    <span class="tree-chevron">${CHEVRON_SVG}</span>
    <span class="tree-phase-name">${phase}</span>
    <div class="tree-phase-meta">
      <span class="tree-chip tc-total">${total}건</span>
      ${done  > 0 ? `<span class="tree-chip tc-done">완료 ${done}</span>` : ''}
      ${delay > 0 ? `<span class="tree-chip tc-delay">지연 ${delay}</span>` : ''}
      ${prog  > 0 ? `<span class="tree-chip tc-prog">진행 ${prog}</span>` : ''}
      <div class="tree-prog-bar"><div class="tree-prog-fill" style="width:${avgProg}%"></div></div>
      <span class="tree-prog-pct">${avgProg}%</span>
    </div>
  </div>
  <div class="tree-phase-body">`;

        actMap.forEach((tasks, activity) => {
          const aDone  = tasks.filter(t => taskStatus(t) === 'done').length;
          const aDelay = tasks.filter(t => taskStatus(t) === 'delay').length;
          const aProg  = tasks.filter(t => taskStatus(t) === 'progress').length;
          const aAvg   = tasks.length ? Math.round(tasks.reduce((s,t) => s + t.progress, 0) / tasks.length) : 0;
          const actOpen = hasFilter ? 'open' : '';

          html += `
    <div class="tree-activity ${actOpen}">
      <div class="tree-act-hd" onclick="this.closest('.tree-activity').classList.toggle('open')">
        <span class="tree-chevron">${CHEVRON_SVG}</span>
        <span class="tree-act-name">${activity}</span>
        <div class="tree-act-meta">
          <span class="tree-chip tc-total" style="font-size:10.5px">${tasks.length}건</span>
          ${aDelay > 0 ? `<span class="tree-chip tc-delay" style="font-size:10.5px">지연 ${aDelay}</span>` : ''}
          ${aProg  > 0 ? `<span class="tree-chip tc-prog" style="font-size:10.5px">진행 ${aProg}</span>` : ''}
          <div class="tree-prog-bar" style="width:60px"><div class="tree-prog-fill" style="width:${aAvg}%"></div></div>
          <span class="tree-prog-pct">${aAvg}%</span>
        </div>
      </div>
      <div class="tree-act-body">`;

          tasks.forEach(t => {
            const st = taskStatus(t);
            const dd = delayDays(t);
            const start = parseDate(t.start), end = parseDate(t.end);
            const actualEnd = parseDate(t.actual_end);
            const dateStr = `${fmt(start)} ~ ${fmt(end)}${st === 'done' && actualEnd ? ` · 완료 ${fmt(actualEnd)}` : ''}`;
            const delayText = dd > 0
              ? `<span class="t-delaydays" style="font-size:10px;padding:1px 6px;flex-shrink:0">${dd}일 지연</span>`
              : (st === 'delay' && isDelayByDate(t) ? `<span class="t-delaydays" style="font-size:10px;padding:1px 6px;flex-shrink:0">마감일 도래</span>` : '');
            const rowCls = st === 'delay' ? 'is-delay' : (st === 'done' ? 'is-done' : '');
            const barColor = st === 'done' ? 'background:var(--green)' : (st === 'delay' ? 'background:var(--red)' : '');
            const ap = t.actualProgress;
            const showDual = ap !== undefined && ap !== t.progress;
            const titleAttr = t.deliverable ? ` title="${t.task}${t.deliverable ? ' | 산출물: ' + t.deliverable : ''}"` : '';
            const pctHtml = showDual
              ? `<span class="tree-task-pct">${t.progress}%·<span style="color:var(--green)">${ap}%</span></span>`
              : `<span class="tree-task-pct">${t.progress}%</span>`;

            html += `
        <div class="tree-task-row ${rowCls}">
          <div style="display:flex;align-items:center">${badgeHtml(st)}</div>
          <div class="tree-task-name"${titleAttr}><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${t.task}</span>${delayText}</div>
          <div class="tree-task-dates">${dateStr}</div>
          <div class="tree-task-bar">
            <div class="mini-bar-bg" style="flex:1"><div class="mini-bar-fill" style="width:${t.progress}%;${barColor}"></div></div>
            ${pctHtml}
          </div>
        </div>`;
          });

          html += `</div></div>`;
        });

        html += `</div></div>`;
      });

      document.getElementById('wbsTree').innerHTML = html;
    }

    function setTreeAll(open) {
      const tree = document.getElementById('wbsTree');
      tree.querySelectorAll('.tree-phase, .tree-activity').forEach(el => {
        open ? el.classList.add('open') : el.classList.remove('open');
      });
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

    function renderPhaseSummary() {
      const grouped = groupByPhase(WBS_DATA);
      let html = '';
      grouped.forEach((tasks, phase) => {
        const total = tasks.length;
        const doneCnt = tasks.filter(t => taskStatus(t) === 'done').length;
        const delayCnt = tasks.filter(t => taskStatus(t) === 'delay').length;
        const progCnt = tasks.filter(t => taskStatus(t) === 'progress').length;
        const totalW = tasks.reduce((s, t) => s + (t.weight || 0), 0) || 1;
        const avgProgress = Math.round(tasks.reduce((s, t) => s + (t.weight || 0) * t.progress, 0) / totalW);
        html += `
    <div class="phase-group">
      <div class="phase-title">${phase} <span style="color:var(--text-sub);font-weight:500;">(${total}건)</span></div>
      <div class="task-row ${delayCnt > 0 ? 'is-delay' : ''}">
        <div class="task-info">
          <div class="t-meta">완료 ${doneCnt} · 진행중 ${progCnt} · <span style="${delayCnt > 0 ? 'color:var(--red);font-weight:700;' : ''}">지연 ${delayCnt}</span></div>
        </div>
        <div class="task-bar-wrap" style="width:240px;">
          <div class="task-bar-bg ${delayCnt > 0 ? 'delaybar' : ''}"><div class="task-bar-fill" style="width:${avgProgress}%;"></div></div>
          <div class="task-pct">가중 평균 공정율 ${avgProgress}%</div>
        </div>
      </div>
    </div>`;
      });
      document.getElementById('phaseSummary').innerHTML = html;
    }

    function renderAll() {
      const [wMon, wSun] = getWeekRange(TODAY);
      const [nMon, nSun] = getWeekRange(addDays(TODAY, 7));
      const nFri = addDays(nMon, 4);

      document.getElementById('todayLabel').textContent = fmt(TODAY) + ' (' + ['일', '월', '화', '수', '목', '금', '토'][TODAY.getDay()] + ')';
      document.getElementById('srcTag').textContent = '데이터 출처: ' + CURRENT_SOURCE_NAME;

      const projStart = parseDate('2026-05-28');
      const projEnd = parseDate('2027-01-29');
      const dday = Math.round((projEnd - TODAY) / 86400000);
      document.getElementById('ddayLabel').textContent = dday >= 0 ? `사업종료 D-${dday}` : `사업종료 D+${-dday}`;

      const thisLabel = `${fmtShort(wMon)} ~ ${fmtShort(wSun)}`;
      const nextLabel = `${fmtShort(nMon)} ~ ${fmtShort(nFri)}`;
      ['thisWeekRangeLabel', 'thisWeekRangeLabel2'].forEach(id => document.getElementById(id).textContent = thisLabel);
      ['nextWeekRangeLabel', 'nextWeekRangeLabel2'].forEach(id => document.getElementById(id).textContent = nextLabel);

      const thisWeekTasks = getThisWeekWorkTasks(wMon, wSun);     // 금주: 완료·진행·지연 업무 전체
      const nextWeekTasks = getTasksInRange(nMon, nFri);          // 차주: 월~금 예정 업무
      const delayTasks = getDelayTasks();

      renderTaskGroup(document.getElementById('thisWeekList'), thisWeekTasks, '이번 주에 진행한 TASK가 없습니다.', false, true);
      renderTaskGroup(document.getElementById('nextWeekList'), nextWeekTasks, '다음 주에 예정된 TASK가 없습니다.');
      renderTaskGroup(document.getElementById('delayList'), delayTasks, '현재 지연된 TASK가 없습니다. 계획대로 진행 중입니다.', true);

      renderDashboardPreview(document.getElementById('dashThisPreview'), thisWeekTasks, 4, '이번 주에 진행한 TASK가 없습니다.');
      renderDashboardPreview(document.getElementById('dashNextPreview'), nextWeekTasks, 4, '다음 주에 예정된 TASK가 없습니다.');
      renderDashboardPreview(document.getElementById('dashDelayPreview'), delayTasks, 4, '현재 지연된 TASK가 없습니다. 계획대로 진행 중입니다.', (a, b) => delayDays(b) - delayDays(a));

      // 사이드바 카운트
      document.getElementById('navCountThis').textContent = thisWeekTasks.length;
      document.getElementById('navCountNext').textContent = nextWeekTasks.length;
      document.getElementById('navCountDelay').textContent = delayTasks.length;
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

      // 공정율 게이지
      const planRate = calcPlanRate();
      const actualRate = calcActualRate();
      const gap = actualRate - planRate;
      document.getElementById('phPlanRate').textContent = planRate.toFixed(1) + '%';
      document.getElementById('phActualRate').textContent = actualRate.toFixed(1) + '%';
      document.getElementById('phGapRate').textContent = (gap >= 0 ? '+' : '') + gap.toFixed(1) + '%p';
      document.getElementById('phGapRate').style.color = gap < -0.5 ? '#B91C1C' : '#047857';
      document.getElementById('phDelayCnt').textContent = delayTasks.length + '건';
      document.getElementById('phMainLabel').textContent = gap < -0.5 ? '계획 대비 지연 중' : (gap > 0.5 ? '계획 대비 선행 중' : '계획대로 진행 중');
      document.getElementById('gaugeLabel').innerHTML = actualRate.toFixed(1) + '%<span class="g-sub">실적</span>';
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
      document.body.classList.toggle('sidebar-collapsed');
    }

    /* ===== 관리자 인증 ===== */
    const PROTECTED_PAGES = new Set(['p-members', 'p-mail']);
    let _authed = false;
    let _pendingPage = '';

    function requireAuth(pageId) {
      if (!PROTECTED_PAGES.has(pageId) || _authed) { navigateTo(pageId); return; }
      _pendingPage = pageId;
      document.getElementById('authOverlay').classList.add('open');
      const inp = document.getElementById('authPwInput');
      inp.type = 'password';
      inp.value = '';
      const error = document.getElementById('authError');
      error.textContent = '';
      error.classList.remove('show');
      setTimeout(() => inp.focus(), 80);
    }

    function submitAuth() {
      const val = document.getElementById('authPwInput').value;
      if (val === '4343') {
        _authed = true;
        closeAuthModal();
        navigateTo(_pendingPage);
      } else {
        const error = document.getElementById('authError');
        error.textContent = '비밀번호가 올바르지 않습니다. 다시 입력해 주세요.';
        error.classList.add('show');
        const inp = document.getElementById('authPwInput');
        inp.value = '';
        inp.focus();
      }
    }

    function closeAuthModal() {
      document.getElementById('authOverlay').classList.remove('open');
    }

    function toggleAuthPassword() {
      const input = document.getElementById('authPwInput');
      const button = document.querySelector('.auth-password-toggle');
      const show = input.type === 'password';
      input.type = show ? 'text' : 'password';
      button.setAttribute('aria-pressed', String(show));
      button.setAttribute('aria-label', show ? '비밀번호 숨기기' : '비밀번호 표시');
      input.focus();
    }

    document.addEventListener('keydown', event => {
      if (event.key === 'Escape' && document.getElementById('authOverlay').classList.contains('open')) closeAuthModal();
    });

    function navigateTo(pageId) {
      document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
      document.querySelectorAll('.page').forEach(p => p.classList.remove('on'));
      const nav = document.querySelector(`.nav-item[data-page="${pageId}"]`);
      if (nav) nav.classList.add('active');
      const page = document.getElementById(pageId);
      if (page) page.classList.add('on');
      if (pageId === 'p-history') renderHistoryPage();
      if (pageId === 'p-members') renderMembersPage();
      if (pageId === 'p-mail') renderMailPage();
    }

    document.querySelectorAll('.nav-item').forEach(item => {
      item.addEventListener('click', () => requireAuth(item.dataset.page));
    });

    function populateFileSelect(activeId) {
      getCombinedHistoryList().then(all => {
        const sel = document.getElementById('fileSelect');
        if (!sel) return;
        const dates = [...new Set(all.map(r => r.wbsDate))].sort((a, b) => b.localeCompare(a));
        sel.innerHTML = '<option value="">날짜 선택</option>' +
          dates.map(d => {
            const rec = all.find(r => r.wbsDate === d);
            return `<option value="${rec.id}"${rec.id === activeId ? ' selected' : ''}>${d}</option>`;
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
        select.innerHTML = '<option value="current">기준일 선택 — 현재 분석 (' + fmt(TODAY) + ')</option>' +
          records.map(r => `<option value="${r.id}">${r.wbsDate} · ${r.fileName}</option>`).join('');
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
      '금주 지연 업무와 차주 할일에 대해서 안내드립니다.\n\n' +
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
        setMailBodyWithNotice(editor, text, `기준일 ${fmt(TODAY)}의 업무 텍스트를 본문에 넣었습니다.`);
        return;
      }
      fetchFullRecord(Number(value)).then(record => {
        if (!record) { showToast('선택한 분석 결과를 찾을 수 없습니다.', 'err'); return; }
        const baseline = startOfDay(parseDate(record.wbsDate));
        const text = buildNextWeekText(record.tasks, baseline, record.fileName);
        if (!text) { showToast('선택한 기준일에 복사할 업무가 없습니다.', ''); return; }
        setMailBodyWithNotice(editor, text, `기준일 ${record.wbsDate}의 업무 텍스트를 본문에 넣었습니다.`);
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
          const msg = `<span class="mail-chip-empty">등록된 구성원이 없습니다. <a onclick="navigateTo('p-members')">구성원 관리</a>에서 먼저 추가하세요.</span>`;
          toWrap.innerHTML = msg;
          ccWrap.innerHTML = msg;
          return;
        }

        toWrap.innerHTML = members.map(m =>
          `<span class="mail-chip" id="to-${m.id}" onclick="toggleMailChip(${m.id},'to')">
            ${m.name} <span class="mail-chip-sub">${m.role}</span>
          </span>`
        ).join('');

        ccWrap.innerHTML = members.map(m =>
          `<span class="mail-chip" id="cc-${m.id}" onclick="toggleMailChip(${m.id},'cc')">
            ${m.name} <span class="mail-chip-sub">${m.role}</span>
          </span>`
        ).join('');
      });
    }

    function toggleMailChip(id, area) {
      const set = area === 'to' ? _mailTo : _mailCc;
      const el = document.getElementById(`${area}-${id}`);
      if (!el) return;
      if (set.has(id)) { set.delete(id); el.classList.remove('selected'); }
      else             { set.add(id);    el.classList.add('selected'); }
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
      document.getElementById('ejsPanel').classList.remove('open');
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
        badge.className = 'ejs-status-badge ' + (ok ? 'ok' : 'warn');
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
    function toggleEjsPanel() {
      document.getElementById('ejsPanel').classList.toggle('open');
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
      btn.innerHTML = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> 발송 중...';

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
      document.querySelectorAll('#mailToChips .mail-chip, #mailCcChips .mail-chip')
        .forEach(el => el.classList.remove('selected'));
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
      if (tbody.children.length === 1 && tbody.children[0].children.length === 1) {
        tbody.innerHTML = '';
      }
      const now = new Date();
      const tstr = fmt(startOfDay(now)) + ' ' + ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${tstr}</td><td>${filename}</td><td>${count}건</td><td>${ok ? '<span style="color:var(--green);font-weight:700;">적용 완료</span>' : '<span style="color:var(--red);font-weight:700;">실패</span>'}</td>`;
      tbody.prepend(tr);
    }

    function handleFile(file) {
      const statusEl = document.getElementById('uploadStatus');
      statusEl.className = 'upload-status show';
      statusEl.textContent = `"${file.name}" 분석 중...`;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const buffer = e.target.result;
          const data = new Uint8Array(buffer);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const { tasks: parsed, todayStr } = parseWbsWorkbook(workbook);
          if (parsed.length === 0) {
            throw new Error('인식된 TASK가 없습니다. "WBS(계획)" 시트 구조를 확인해주세요.');
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
          const dateNote = fileDateStr ? ` · 기준일 ${fileDateStr} (파일명)` : (todayStr ? ` · 기준일 ${todayStr} (엑셀 Today 셀)` : '');
          statusEl.className = 'upload-status show ok';
          statusEl.textContent = `✔ "${file.name}" 적용 완료 — TASK ${parsed.length}건 인식${dateNote}`;
          addUploadHistory(file.name, parsed.length, true);
          showSaveBar(file.name, fmt(TODAY), parsed.length);
        } catch (err) {
          statusEl.className = 'upload-status show err';
          statusEl.textContent = `⚠ 파일을 분석하지 못했습니다: ${err.message}`;
          addUploadHistory(file.name, 0, false);
        }
      };
      reader.onerror = function () {
        statusEl.className = 'upload-status show err';
        statusEl.textContent = '⚠ 파일을 읽는 중 오류가 발생했습니다.';
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
    let _toastTimer = null;
    function showToast(msg, type) {
      const el = document.getElementById('toast');
      el.textContent = msg;
      el.className = 'show' + (type ? ' ' + type : '');
      clearTimeout(_toastTimer);
      _toastTimer = setTimeout(() => { el.className = ''; }, 3200);
    }

    /* ===================== 저장 관련 상태 ===================== */
    let CURRENT_FILE_BUFFER = null;

    function showSaveBar(fileName, wbsDate, taskCount) {
      const bar = document.getElementById('saveBar');
      document.getElementById('saveBarTitle').textContent = `"${fileName}" — ${taskCount}건 인식, 기준일 ${wbsDate}`;
      document.getElementById('saveBarSub').textContent = '분석 결과를 확인한 후 최종 저장하면 조회 메뉴에서 다시 불러올 수 있습니다.';
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
              showToast('✔ 클라우드 및 로컬 저장 완료!', 'ok');
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
              showToast('✔ 로컬 저장 완료 (클라우드 미연동)', 'ok');
              updateHistoryCount();
              populateFileSelect(id);
              renderHistoryPage();
              updateUrlParam(id);
            }).catch(err => showToast('저장 오류: ' + err.message, 'err'));
          } else {
            dbAdd(record).then(id => {
              showToast('✔ 로컬 저장 완료 (클라우드 미연동)', 'ok');
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
        const label = document.getElementById('historyCountLabel');
        if (list.length > 0) {
          badge.textContent = list.length;
          badge.style.display = 'inline-flex';
        } else {
          badge.style.display = 'none';
        }
        if (label) label.textContent = `총 ${list.length}건`;
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
        badge.textContent = '연동 됨';
        badge.className = 'ejs-status-badge ok';
      } else {
        badge.textContent = '연동 안 됨';
        badge.className = 'ejs-status-badge warn';
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
      renderHistoryPage();
    }

    /* ===================== 프로젝트 구성원 관리 ===================== */
    function toggleMemberForm(editData) {
      const panel = document.getElementById('memberFormPanel');
      if (editData) {
        document.getElementById('memberFormTitle').textContent = '구성원 수정';
        document.getElementById('mfEditId').value = editData.id;
        document.getElementById('mfCreatedAt').value = editData.createdAt || '';
        document.getElementById('mfName').value = editData.name || '';
        document.getElementById('mfRole').value = editData.role || '';
        document.getElementById('mfEmail').value = editData.email || '';
        document.getElementById('mfPhone').value = editData.phone || '';
        document.getElementById('mfTitle').value = editData.title || '';
        panel.classList.add('open');
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      } else {
        if (panel.classList.contains('open')) {
          closeMemberForm();
        } else {
          document.getElementById('memberFormTitle').textContent = '구성원 추가';
          document.getElementById('mfEditId').value = '';
          document.getElementById('mfCreatedAt').value = '';
          document.getElementById('mfName').value = '';
          document.getElementById('mfRole').value = '';
          document.getElementById('mfEmail').value = '';
          document.getElementById('mfPhone').value = '';
          document.getElementById('mfTitle').value = '';
          panel.classList.add('open');
          document.getElementById('mfName').focus();
        }
      }
    }

    function closeMemberForm() {
      document.getElementById('memberFormPanel').classList.remove('open');
    }

    function saveMember() {
      const fields = [
        { id: 'mfName',  label: '이름' },
        { id: 'mfRole',  label: '역할' },
        { id: 'mfEmail', label: '이메일' },
        { id: 'mfPhone', label: '연락처' },
        { id: 'mfTitle', label: '직책' },
      ];
      for (const f of fields) {
        const el = document.getElementById(f.id);
        if (!el.value.trim()) {
          showToast(`${f.label}을(를) 입력하세요.`, 'err');
          el.focus();
          return;
        }
      }

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
        if (navBadge) { navBadge.textContent = count; navBadge.style.display = count > 0 ? 'inline-flex' : 'none'; }

        if (count === 0) {
          tbody.innerHTML = `<tr><td colspan="7"><div class="member-empty">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
              <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/>
              <path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>
            </svg>
            등록된 구성원이 없습니다.<br>위의 <b>구성원 추가</b> 버튼을 눌러 추가하세요.
          </div></td></tr>`;
          return;
        }

        const dash = '<span style="color:var(--text-muted)">-</span>';
        tbody.innerHTML = all.map((m, i) => {
          const role = m.role ? `<span class="member-role-badge dev">${m.role}</span>` : dash;
          return `<tr>
            <td style="color:var(--text-muted);font-size:12px">${i + 1}</td>
            <td><strong>${m.name}</strong></td>
            <td>${role}</td>
            <td>${m.title || dash}</td>
            <td>${m.email}</td>
            <td>${m.phone || dash}</td>
            <td><div class="member-actions">
              <button class="btn-m-edit" onclick="editMember(${m.id})">수정</button>
              <button class="btn-m-del" onclick="deleteMember(${m.id})">삭제</button>
            </div></td>
          </tr>`;
        }).join('');
      });
    }

    function fmtSavedAt(iso) {
      const d = new Date(iso);
      return `${d.getFullYear()}-${('0'+(d.getMonth()+1)).slice(-2)}-${('0'+d.getDate()).slice(-2)} ${('0'+d.getHours()).slice(-2)}:${('0'+d.getMinutes()).slice(-2)}`;
    }

    // 분석 이력은 IndexedDB/클라우드에서 비동기로 불러온다. 이전 요청이 늦게
    // 끝나면 방금 고른 날짜의 결과를 다시 전체 목록으로 덮어쓰지 않도록 한다.
    let historyRenderVersion = 0;

    function renderHistoryPage() {
      const renderVersion = ++historyRenderVersion;
      const sel = document.getElementById('histSearchName');
      if (!sel) return;
      const selectedDate = sel.value;
      const container = document.getElementById('historyList');
      const label = document.getElementById('historyCountLabel');

      getCombinedHistoryList().then(all => {
        if (renderVersion !== historyRenderVersion) return;
        if (label) label.textContent = `총 ${all.length}건`;
        const badge = document.getElementById('navCountHistory');
        if (badge) {
          if (all.length > 0) { badge.textContent = all.length; badge.style.display = 'inline-flex'; }
          else badge.style.display = 'none';
        }

        // SELECT 옵션 갱신 (wbsDate 오름차순, 중복 제거)
        const dates = [...new Set(all.map(r => r.wbsDate))].sort((a, b) => a.localeCompare(b));
        sel.innerHTML = '<option value="">전체 날짜 보기</option>' +
          dates.map(d => `<option value="${d}"${d === selectedDate ? ' selected' : ''}>${d}</option>`).join('');

        const list = all.filter(r => !selectedDate || r.wbsDate === selectedDate)
          .sort((a, b) => a.wbsDate.localeCompare(b.wbsDate));

        if (list.length === 0) {
          container.innerHTML = all.length === 0
            ? `<div class="hist-empty"><svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 8v4l3 3M3.05 11a9 9 0 1 0 .5-3M3 4v4h4"/></svg>저장된 분석 결과가 없습니다.<br>WBS 업데이트 후 <b>최종 저장</b>을 눌러 기록을 남기세요.</div>`
            : `<div class="hist-empty">검색 조건에 해당하는 결과가 없습니다.</div>`;
          return;
        }

        const currentWbsDate = fmt(TODAY);
        let html = '<div class="hist-grid">';
        list.forEach(r => {
          const isActive = r.wbsDate === currentWbsDate && r.fileName === CURRENT_SOURCE_NAME;
          const s = r.summary;
          const gap = s.actualRate - s.planRate;
          const gapStr = (gap >= 0 ? '+' : '') + gap.toFixed(1) + '%p';
          const gapColor = gap < -0.5 ? 'c-red' : (gap > 0.5 ? 'c-green' : '');
          html += `
<div class="hist-card${isActive ? ' active-snapshot' : ''}">
  <div class="hist-date-block">
    <div class="hist-wbs-date">${r.wbsDate}<small>기준일</small></div>
    <span class="hist-active-badge">현재 분석 중</span>
  </div>
  <div class="hist-file-block">
    <div class="hist-filename" title="${r.fileName}">📎 ${r.fileName}</div>
    <div class="hist-saved-at">저장: ${fmtSavedAt(r.savedAt)}</div>
  </div>
  <div class="hist-stats">
    <div class="hs-item"><div class="hs-label">전체 TASK</div><div class="hs-value">${s.total}</div></div>
    <div class="hs-item"><div class="hs-label">완료</div><div class="hs-value c-green">${s.done}</div></div>
    <div class="hs-item"><div class="hs-label">지연</div><div class="hs-value c-red">${s.delay}</div></div>
    <div class="hs-item"><div class="hs-label">계획 공정율</div><div class="hs-value">${s.planRate}%</div></div>
    <div class="hs-item"><div class="hs-label">실적 공정율</div><div class="hs-value c-accent">${s.actualRate}%</div></div>
    <div class="hs-item"><div class="hs-label">Gap</div><div class="hs-value ${gapColor}">${gapStr}</div></div>
  </div>
  <div class="hist-footer">
    <div class="hist-actions">
      <button class="btn-restore" onclick="restoreAnalysis(${r.id})">불러오기</button>
      <button class="btn-link" onclick="copyAnalysisLink(${r.id})">링크 복사</button>
      <button class="btn-download" onclick="downloadAnalysisFile(${r.id})">파일 다운로드</button>
      <button class="btn-del" onclick="deleteAnalysis(${r.id})">삭제</button>
    </div>
  </div>
</div>`;
        });
        html += '</div>';
        container.innerHTML = html;
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
      const delayTasks = data.filter(t => !isDone(t) && (isDelayByBaseline(t) || (isInProgress(t) && hasProgressGap(t))))
        .sort((a, b) => taskDelayDays(b) - taskDelayDays(a) || ((b.progress - (b.actualProgress ?? b.progress)) - (a.progress - (a.actualProgress ?? a.progress))));
      const [nMon, nSun] = getWeekRange(addDays(baseline, 7));
      const nFri = addDays(nMon, 4);
      const tasks = data.filter(t => {
        if (isDone(t)) return false;
        const start = parseDate(t.start), end = parseDate(t.end);
        return start && end && rangesOverlap(start, end, nMon, nFri);
      });
      if (tasks.length === 0 && delayTasks.length === 0) return '';

      const rangeStr = `${fmt(nMon)} ~ ${fmt(nFri)}`;
      const lines = [];
      lines.push(`[금주의 지연 업무] 기준일 ${fmt(baseline)}`);
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
      lines.push(`[차주 할 일] ${rangeStr}`);
      lines.push('');

      if (tasks.length === 0) {
        lines.push('다음 주에 예정된 TASK가 없습니다.');
      } else {
        let lastPhase = '', lastAct = '';
        tasks.forEach((t, i) => {
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
      lines.push(`※ 기준일: ${fmt(baseline)} / 출처: ${sourceName}`);
      return lines.join('\n');
    }

    function copyNextWeekText() {
      const text = buildNextWeekText(WBS_DATA, TODAY, CURRENT_SOURCE_NAME);
      if (!text) { showToast('복사할 차주 할 일 또는 지연 업무가 없습니다.', ''); return; }
      const [nMon] = getWeekRange(addDays(TODAY, 7));
      const nFri = addDays(nMon, 4);
      const tasks = getTasksInRange(nMon, nFri);
      const delayTasks = getDelayTasks();

      navigator.clipboard.writeText(text)
        .then(() => showToast(`지연 업무 ${delayTasks.length}건 · 차주 할 일 ${tasks.length}건 복사 완료`, 'ok'))
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
      showToast(`✔ "${r.fileName}" (기준일 ${r.wbsDate}) 불러오기 완료`, 'ok');
      updateUrlParam(id);
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

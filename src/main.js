import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import './styles.css';

const DAYS = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
const PERIODS = ['1/2', '3/4', '5/6', '7/8', '9/10'];
const PERIOD_DETAILS = [
  { index: 1, start: '08:00', end: '08:50' },
  { index: 2, start: '09:00', end: '09:50' },
  { index: 3, start: '10:10', end: '11:00' },
  { index: 4, start: '11:10', end: '12:00' },
  { index: 5, start: '13:30', end: '14:20' },
  { index: 6, start: '14:30', end: '15:20' },
  { index: 7, start: '15:40', end: '16:30' },
  { index: 8, start: '16:40', end: '17:30' },
  { index: 9, start: '18:30', end: '19:20' },
  { index: 10, start: '19:30', end: '20:20' }
];
const TOTAL_SHEET_NAMES = new Set(['总表', '总课表', '汇总']);
const DB_NAME = 'course-timetable-db';
const DB_VERSION = 1;
const STORE_NAME = 'settings';
const APP_STATE_KEY = 'main';

const MAJOR_ALIASES = {
  CS: '计算机系统',
  AI: '人工智能',
  SE: '软件工程',
  TH: '天河班',
  NE: '网络工程',
  IS: '信息安全'
};

const state = {
  view: 'schedule',
  schedules: [],
  activeScheduleId: '',
  scheduleName: '',
  sourceName: '',
  timetable: null,
  selectedMajor: '',
  selectedWeek: '',
  selectedCourseKeys: new Set(),
  showWeekend: true,
  occurrenceEdits: { deleted: [], roomOverrides: {}, added: [] },
  setupMode: false,
  status: 'loading',
  saveStatus: '',
  courseSearch: '',
  activeConflictKey: '',
  activeEditCellKey: '',
  touchStartX: 0
};

const app = document.querySelector('#app');

init();

async function init() {
  render();
  registerServiceWorker();
  const savedState = await loadSavedState();
  if (savedState) {
    state.schedules = normalizeSavedSchedules(savedState);
    state.activeScheduleId = savedState.activeScheduleId ?? state.schedules[0]?.id ?? '';
    const active = state.schedules.find((item) => item.id === state.activeScheduleId) ?? state.schedules[0];
    if (active) applySchedule(active);
  }
  state.status = 'ready';
  render();
}

function render() {
  if (state.status === 'loading') {
    app.innerHTML = `<main class="appShell"><section class="emptyState"><h2>正在读取本地课表</h2><p>应用会加载上一次保存的课表。</p></section></main>`;
    return;
  }

  ensureDefaultSelections();
  const events = getVisibleEvents();
  const catalog = getCourseCatalog();

  app.innerHTML = `
    <main class="appShell">
      ${state.view === 'schedule' ? renderSchedulePage(events, catalog) : renderLibraryPage(catalog)}
      ${renderBottomNav()}
      ${renderConflictModal(events)}
      ${renderCellEditModal(events, catalog)}
    </main>
  `;
  bindEvents();
}

function ensureDefaultSelections() {
  const majors = state.timetable?.majors ?? [];
  const today = getCurrentAcademicDay();
  const weeks = getAvailableWeeks();
  if (!state.selectedMajor && majors.length) state.selectedMajor = majors[0].id;
  if (!state.selectedCourseKeys.size && state.selectedMajor) {
    state.selectedCourseKeys = new Set(getCourseKeysForMajor(state.selectedMajor));
  }
  if (!state.selectedWeek) state.selectedWeek = today?.week ?? weeks[0] ?? '';
}

function renderSchedulePage(events, catalog) {
  if (!state.timetable) {
    return `
      <section class="pageHeader">
        <div><p class="eyebrow">日常课表</p><h1>暂无已保存课表</h1></div>
      </section>
      <section class="emptyState">
        <h2>请先导入并保存课表</h2>
        <p>进入底部“管理”页面导入 CSV 或 Excel，并完成专业和课程选择。</p>
      </section>
    `;
  }

  const schedules = state.schedules;
  const weeks = getAvailableWeeks();
  const today = getCurrentAcademicDay();
  const todayLabel = formatDate(new Date());
  const currentWeekText = today ? `当前为第 ${today.week} 周` : '未匹配到当前日期';
  const activeMajor = state.timetable.majors.find((item) => item.id === state.selectedMajor);
  const conflicts = getSelectedConflicts();

  return `
    <section class="pageHeader compactHeader">
      <div>
        <h1>${escapeHtml(todayLabel)}</h1>
        <p>${escapeHtml(state.selectedWeek ? `第 ${state.selectedWeek} 周` : '未选择周次')}　${escapeHtml(currentWeekText)}</p>
      </div>
      <div class="quickActions">
        <button id="todayButton" class="iconButton" type="button" ${today ? '' : 'disabled'} title="回到今天">今</button>
        <button class="iconButton" type="button" data-view="library" title="管理课表">＋</button>
      </div>
    </section>

    <section class="displayControls">
      <label>
        <span>课表</span>
        <select id="displayScheduleSelect">
          ${schedules.map((schedule) => `<option value="${escapeHtml(schedule.id)}" ${schedule.id === state.activeScheduleId ? 'selected' : ''}>${escapeHtml(schedule.name)}</option>`).join('')}
        </select>
      </label>
      <label>
        <span>周次</span>
        <select id="weekSelect">
          ${weeks.map((week) => `<option value="${escapeHtml(week)}" ${week === state.selectedWeek ? 'selected' : ''}>第 ${escapeHtml(week)} 周</option>`).join('')}
        </select>
      </label>
      <label class="inlineCheck"><input id="weekendToggle" type="checkbox" ${state.showWeekend ? 'checked' : ''} /> 周末</label>
    </section>

    ${conflicts.length ? renderConflictSummary(conflicts) : ''}
    ${renderSchedule(activeMajor, events, today)}
    <p class="swipeHint">左右滑动课表切换周数</p>
  `;
}

function renderLibraryPage(catalog) {
  return `
    <section class="pageHeader">
      <div>
        <p class="eyebrow">课表管理</p>
        <h1>已保存课表</h1>
      </div>
      <label class="fileButton">
        <input id="fileInput" type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
        <span>导入新课表</span>
      </label>
    </section>
    ${renderCsvNotice()}
    ${renderScheduleList()}
    ${state.setupMode && state.timetable ? renderSetupPanel(catalog) : ''}
  `;
}

function renderScheduleList() {
  if (!state.schedules.length) {
    return `<section class="emptyState"><h2>还没有保存课表</h2><p>导入 CSV 或 Excel 后，选择专业和课程，再保存为课表档案。</p></section>`;
  }
  return `
    <section class="scheduleList">
      ${state.schedules.map((schedule) => `
        <article class="scheduleCard ${schedule.id === state.activeScheduleId ? 'active' : ''}">
          <div>
            <strong>${escapeHtml(schedule.name)}</strong>
            <span>${escapeHtml(schedule.sourceName || '')}</span>
          </div>
          <div class="cardActions">
            <button class="ghostButton" type="button" data-use-schedule="${escapeHtml(schedule.id)}">使用</button>
            <button class="ghostButton" type="button" data-edit-schedule="${escapeHtml(schedule.id)}">编辑</button>
            <button class="ghostButton danger" type="button" data-delete-schedule="${escapeHtml(schedule.id)}">删除</button>
          </div>
        </article>
      `).join('')}
    </section>
  `;
}

function renderSetupPanel(catalog) {
  const majors = state.timetable?.majors ?? [];
  const conflicts = getSelectedConflicts();
  return `
    <section class="setupPanel">
      <div class="setupHeader">
        <div><p class="eyebrow">准备课表</p><h2>${escapeHtml(state.sourceName || '新课表')}</h2></div>
        <button id="saveButton" class="primaryButton" type="button">保存课表</button>
      </div>
      <div class="setupFields">
        <label><span>课表名称</span><input id="scheduleNameInput" class="textInput" type="text" value="${escapeHtml(state.scheduleName)}" /></label>
        ${renderMajorDropdown(majors)}
      </div>
      <div class="setupActions">
        <button id="selectMajorButton" class="ghostButton" type="button">全选本专业</button>
        <button id="clearCoursesButton" class="ghostButton" type="button">清空课程</button>
      </div>
      ${renderCoursePicker(catalog, conflicts)}
      ${conflicts.length ? renderConflictSummary(conflicts) : ''}
      ${state.saveStatus ? `<p class="saveStatus">${escapeHtml(state.saveStatus)}</p>` : ''}
    </section>
  `;
}

function renderBottomNav() {
  return `
    <nav class="bottomNav">
      <button class="${state.view === 'schedule' ? 'active' : ''}" type="button" data-view="schedule">课表</button>
      <button class="${state.view === 'library' ? 'active' : ''}" type="button" data-view="library">管理</button>
    </nav>
  `;
}

function renderMajorDropdown(majors) {
  const active = majors.find((major) => major.id === state.selectedMajor) ?? majors[0];
  return `
    <div class="majorField">
      <span>我的专业</span>
      <details class="majorDropdown">
        <summary>${escapeHtml(active?.label ?? '请选择专业')}</summary>
        <div class="majorMenu">
          ${majors.map((major) => `
            <button class="${major.id === state.selectedMajor ? 'active' : ''}" type="button" data-major-option="${escapeHtml(major.id)}">
              ${escapeHtml(major.label)}
            </button>
          `).join('')}
        </div>
      </details>
    </div>
  `;
}

function bindEvents() {
  document.querySelectorAll('[data-view]').forEach((button) => {
    button.addEventListener('click', async () => {
      await autoSaveIfReady();
      state.view = button.dataset.view;
      render();
    });
  });
  document.querySelector('#fileInput')?.addEventListener('change', onFileChange);
  document.querySelector('#displayScheduleSelect')?.addEventListener('change', async (event) => switchSchedule(event.target.value));
  document.querySelectorAll('[data-use-schedule]').forEach((button) => button.addEventListener('click', async () => {
    await switchSchedule(button.dataset.useSchedule);
    state.view = 'schedule';
    render();
  }));
  document.querySelectorAll('[data-edit-schedule]').forEach((button) => button.addEventListener('click', async () => {
    await switchSchedule(button.dataset.editSchedule);
    state.view = 'library';
    state.setupMode = true;
    render();
  }));
  document.querySelectorAll('[data-delete-schedule]').forEach((button) => button.addEventListener('click', () => deleteSchedule(button.dataset.deleteSchedule)));
  document.querySelector('#scheduleNameInput')?.addEventListener('change', async (event) => {
    state.scheduleName = normalize(event.target.value) || state.sourceName || '未命名课表';
    await persistCurrentScheduleDraft();
    render();
  });
  document.querySelectorAll('[data-major-option]').forEach((button) => button.addEventListener('click', () => {
    state.selectedMajor = button.dataset.majorOption;
    state.selectedCourseKeys = new Set(getCourseKeysForMajor(state.selectedMajor));
    state.selectedWeek = '';
    state.saveStatus = '';
    render();
  }));
  document.querySelector('#weekSelect')?.addEventListener('change', (event) => {
    state.selectedWeek = event.target.value;
    persistCurrentScheduleDraft();
    render();
  });
  document.querySelector('#todayButton')?.addEventListener('click', () => {
    const today = getCurrentAcademicDay();
    if (today) {
      state.selectedWeek = today.week;
      persistCurrentScheduleDraft();
      render();
    }
  });
  document.querySelector('#weekendToggle')?.addEventListener('change', async (event) => {
    state.showWeekend = event.target.checked;
    await autoSaveIfReady();
    render();
  });
  document.querySelector('#saveButton')?.addEventListener('click', saveCurrentState);
  document.querySelector('#selectMajorButton')?.addEventListener('click', () => {
    state.selectedCourseKeys = new Set(getCourseKeysForMajor(state.selectedMajor));
    state.selectedWeek = '';
    render();
  });
  document.querySelector('#clearCoursesButton')?.addEventListener('click', () => {
    state.selectedCourseKeys = new Set();
    state.selectedWeek = '';
    render();
  });
  document.querySelector('#courseSearch')?.addEventListener('input', (event) => {
    state.courseSearch = event.target.value;
    filterCourseOptions(state.courseSearch);
  });
  document.querySelectorAll('[data-course-key]').forEach((input) => {
    input.addEventListener('change', () => {
      const key = input.dataset.courseKey;
      if (input.checked) state.selectedCourseKeys.add(key);
      else state.selectedCourseKeys.delete(key);
      input.closest('.courseOption')?.classList.toggle('selected', input.checked);
    });
  });
  document.querySelector('[data-swipe-week]')?.addEventListener('touchstart', (event) => {
    state.touchStartX = event.changedTouches[0].clientX;
  }, { passive: true });
  document.querySelector('[data-swipe-week]')?.addEventListener('touchend', (event) => {
    const delta = event.changedTouches[0].clientX - state.touchStartX;
    if (Math.abs(delta) > 48) changeWeek(delta < 0 ? 1 : -1);
  }, { passive: true });
  document.querySelectorAll('[data-open-conflict]').forEach((button) => {
    button.addEventListener('click', () => {
      state.activeConflictKey = button.dataset.openConflict;
      state.activeEditCellKey = '';
      render();
    });
  });
  document.querySelectorAll('[data-open-cell]').forEach((button) => {
    button.addEventListener('click', () => {
      state.activeEditCellKey = button.dataset.openCell;
      state.activeConflictKey = '';
      render();
    });
  });
  document.querySelectorAll('[data-close-modal]').forEach((button) => button.addEventListener('click', closeModal));
  document.querySelector('.modalBackdrop')?.addEventListener('click', (event) => {
    if (event.target === event.currentTarget) closeModal();
  });
  document.querySelector('#editRoomForm')?.addEventListener('submit', onEditRoomSubmit);
  document.querySelector('#deleteOccurrenceButton')?.addEventListener('click', onDeleteOccurrence);
  document.querySelector('#addExistingCourseForm')?.addEventListener('submit', onAddExistingCourse);
  document.querySelector('#addCustomCourseForm')?.addEventListener('submit', onAddCustomCourse);
}

async function switchSchedule(scheduleId) {
  await persistCurrentScheduleDraft();
  const next = state.schedules.find((item) => item.id === scheduleId);
  if (next) applySchedule(next);
}

function changeWeek(offset) {
  const weeks = getAvailableWeeks();
  const index = weeks.indexOf(state.selectedWeek);
  const next = weeks[Math.min(Math.max(index + offset, 0), weeks.length - 1)];
  if (next && next !== state.selectedWeek) {
    state.selectedWeek = next;
    persistCurrentScheduleDraft();
    render();
  }
}

function closeModal() {
  state.activeConflictKey = '';
  state.activeEditCellKey = '';
  render();
}

async function onFileChange(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const timetable = file.name.toLowerCase().endsWith('.csv')
      ? await parseCsvFile(file)
      : await parseWorkbookFile(file);
    state.activeScheduleId = createId();
    state.scheduleName = createDefaultScheduleName(file.name);
    state.sourceName = file.name;
    state.timetable = timetable;
    state.selectedMajor = timetable.majors[0]?.id ?? '';
    state.selectedCourseKeys = new Set(getCourseKeysForMajor(state.selectedMajor));
    state.selectedWeek = getCurrentAcademicDay()?.week ?? '';
    state.occurrenceEdits = normalizeOccurrenceEdits();
    state.courseSearch = '';
    state.setupMode = true;
    state.view = 'library';
    state.saveStatus = '请确认专业和课程后保存。';
    persistCurrentSchedule();
    render();
  } catch (error) {
    app.innerHTML = `<main class="appShell"><section class="errorBox"><h1>导入失败</h1><p>${escapeHtml(error.message)}</p><button id="resetButton">重新导入</button></section>${renderBottomNav()}</main>`;
    document.querySelector('#resetButton')?.addEventListener('click', render);
  }
}

async function saveCurrentState() {
  if (!state.timetable) return;
  if (!state.selectedCourseKeys.size) {
    state.saveStatus = '至少选择一门课程后才能保存。';
    render();
    return;
  }
  await saveAppState();
  state.setupMode = false;
  state.view = 'schedule';
  state.saveStatus = '已保存到本地数据库。';
  render();
}

async function deleteSchedule(scheduleId) {
  const target = state.schedules.find((item) => item.id === scheduleId);
  if (!target) return;
  if (!window.confirm(`确定删除“${target.name}”吗？`)) return;
  state.schedules = state.schedules.filter((item) => item.id !== scheduleId);
  if (state.activeScheduleId === scheduleId) {
    const next = state.schedules[0];
    if (next) applySchedule(next);
    else resetCurrentSchedule();
  }
  await saveAppState();
  render();
}

async function autoSaveIfReady() {
  if (!state.timetable || !state.selectedCourseKeys.size) return;
  await persistCurrentScheduleDraft();
}

async function parseWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const titleText = workbook.SheetNames.map((name) => `${name} ${workbook.Sheets[name]?.A1?.v ?? ''}`).join(' ');
  const termYear = inferTermYear(`${file.name} ${titleText}`);
  const sheetNames = workbook.SheetNames;
  const majorSheets = sheetNames.filter((name) => !TOTAL_SHEET_NAMES.has(name));
  const majors = majorSheets.length
    ? majorSheets.map((name) => ({ id: name, label: name }))
    : Object.entries(MAJOR_ALIASES).map(([id, label]) => ({ id, label: `${label} (${id})` }));
  const eventsByMajor = {};
  for (const major of majors) {
    const sheet = workbook.Sheets[major.id];
    eventsByMajor[major.id] = sheet ? parseMajorSheet(sheet, major.id) : parseTotalSheet(workbook, major.id);
  }
  return { type: 'xlsx', termYear, majors, eventsByMajor };
}

function parseMajorSheet(sheet, majorId) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 3) throw new Error(`工作表 ${majorId} 行数不足，无法解析。`);
  const detailStart = rows.findIndex((row) => normalize(row[0]) === '课程名称');
  const scheduleRows = detailStart >= 0 ? rows.slice(2, detailStart) : rows.slice(2);
  const courseDetails = detailStart >= 0 ? parseCourseDetails(rows.slice(detailStart)) : new Map();
  const columns = buildColumnMap(rows[0], rows[1], 2);
  return collectEvents(scheduleRows, columns, courseDetails, 0, 1, majorId);
}

function parseTotalSheet(workbook, majorCode) {
  const totalName = workbook.SheetNames.find((name) => TOTAL_SHEET_NAMES.has(name)) ?? workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[totalName], { header: 1, defval: '' });
  const columns = buildColumnMap(rows[1], rows[2], 3);
  const matchingRows = rows.slice(3).filter((row) => normalize(row[2]) === majorCode || normalize(row[2]) === normalize(MAJOR_ALIASES[majorCode]));
  return collectEvents(matchingRows, columns, new Map(), 0, 1, majorCode);
}

function buildColumnMap(dayRow, slotRow, startCol) {
  const columns = [];
  let currentDay = '';
  const max = Math.max(dayRow.length, slotRow.length);
  for (let col = startCol; col < max; col += 1) {
    const dayCell = normalize(dayRow[col]);
    const slotCell = normalize(slotRow[col]);
    if (dayCell) currentDay = normalizeDay(dayCell);
    if (currentDay && slotCell) columns.push({ col, day: currentDay, slot: normalizeSlot(slotCell) });
  }
  return columns;
}

function collectEvents(rows, columns, details, weekCol, dateCol, majorId) {
  const events = [];
  for (const row of rows) {
    const week = normalize(row[weekCol]);
    const dateRange = normalize(row[dateCol]);
    if (!week || !/^\d+$/.test(week)) continue;
    for (const column of columns) {
      const shortName = normalize(row[column.col]);
      if (!shortName) continue;
      const detail = details.get(shortName) ?? {};
      events.push({
        id: `${majorId}-${week}-${column.day}-${column.slot}-${shortName}`,
        majorId,
        week,
        dateRange,
        day: column.day,
        slot: column.slot,
        title: detail.name ?? shortName,
        shortName,
        teacher: detail.teacher ?? '',
        room: detail.room ?? '',
        hours: detail.hours ?? ''
      });
    }
  }
  return events;
}

function parseCourseDetails(rows) {
  const details = new Map();
  const header = rows[0] ?? [];
  const nameCols = header.map((cell, index) => (normalize(cell) === '课程名称' ? index : -1)).filter((index) => index >= 0);
  for (const nameCol of nameCols) {
    const nextNameCol = nameCols.find((col) => col > nameCol) ?? header.length;
    const block = header.slice(nameCol, nextNameCol);
    const findOffset = (label) => {
      const offset = block.findIndex((cell) => normalize(cell) === label);
      return offset >= 0 ? nameCol + offset : -1;
    };
    const shortCol = findOffset('简称');
    const hoursCol = findOffset('学时');
    const teacherCol = findOffset('授课教员');
    const roomCol = findOffset('授课地点');
    for (const row of rows.slice(1)) {
      const name = normalize(row[nameCol]);
      const shortName = normalize(row[shortCol]);
      if (!name || !shortName) continue;
      details.set(shortName, {
        name,
        shortName,
        hours: normalize(row[hoursCol]),
        teacher: normalize(row[teacherCol]),
        room: normalize(row[roomCol])
      });
    }
  }
  return details;
}

async function parseCsvFile(file) {
  const text = await file.text();
  const result = Papa.parse(text, { header: true, skipEmptyLines: true, transformHeader: normalize });
  if (result.errors.length) throw new Error(result.errors[0].message);
  const majorsById = new Map();
  const eventsByMajor = {};
  for (const row of result.data) {
    const major = pick(row, ['专业', 'major', '班级', 'class']) || '默认';
    const week = pick(row, ['周次', 'week']) || '1';
    const title = pick(row, ['课程名称', '课程', 'course', 'name', 'title']);
    const day = normalizeDay(pick(row, ['星期', 'day', 'weekday']));
    const slot = normalizeSlot(pick(row, ['节次', 'slot', 'period', 'time']));
    if (!title || !day || !slot) continue;
    if (!majorsById.has(major)) majorsById.set(major, { id: major, label: major });
    if (!eventsByMajor[major]) eventsByMajor[major] = [];
    eventsByMajor[major].push({
      id: `${major}-${week}-${day}-${slot}-${title}`,
      majorId: major,
      week,
      dateRange: pick(row, ['日期', 'date', '日期范围', 'dateRange']),
      day,
      slot,
      title,
      shortName: pick(row, ['简称', 'shortName']) || title,
      teacher: pick(row, ['教师', '老师', 'teacher']),
      room: pick(row, ['地点', '教室', 'room', 'location']),
      hours: pick(row, ['学时', 'hours'])
    });
  }
  const majors = Array.from(majorsById.values());
  if (!majors.length) throw new Error('CSV 中没有可识别的课程行。至少需要课程名称、星期、节次字段。');
  return { type: 'csv', termYear: inferTermYear(file.name), majors, eventsByMajor };
}

function getVisibleEvents() {
  const selectedKeys = state.selectedCourseKeys;
  const base = getAllEvents().filter((event) => selectedKeys.has(getCourseKey(event)) && (!state.selectedWeek || event.week === state.selectedWeek));
  return applyOccurrenceEdits(dedupeEventsByCourseAndTime(base), true);
}

function getAvailableWeeks() {
  const selectedKeys = state.selectedCourseKeys;
  const base = selectedKeys.size ? getAllEvents().filter((event) => selectedKeys.has(getCourseKey(event))) : getAllEvents();
  const added = state.occurrenceEdits.added ?? [];
  return Array.from(new Set([...base.map((event) => event.week), ...added.map((event) => event.week)])).sort((a, b) => Number(a) - Number(b));
}

function getAllEvents() {
  return state.timetable ? Object.values(state.timetable.eventsByMajor).flat() : [];
}

function getCourseCatalog() {
  if (!state.timetable) return [];
  const majorsById = new Map(state.timetable.majors.map((major) => [major.id, major.label]));
  const courses = new Map();
  for (const event of getAllEvents()) {
    const key = getCourseKey(event);
    if (!courses.has(key)) {
      courses.set(key, {
        key,
        group: getMajorGroup(event.majorId),
        groupLabel: getMajorGroupLabel(event.majorId),
        title: event.title,
        shortName: event.shortName || event.title,
        teacher: event.teacher,
        room: event.room,
        majors: new Map(),
        weeks: new Set(),
        slots: new Set()
      });
    }
    const course = courses.get(key);
    course.majors.set(event.majorId, majorsById.get(event.majorId) ?? event.majorId);
    course.weeks.add(event.week);
    course.slots.add(`${event.day} ${event.slot}`);
  }
  return Array.from(courses.values()).map((course) => ({
    ...course,
    majorLabels: Array.from(course.majors.values()),
    weekCount: course.weeks.size,
    slotCount: course.slots.size
  })).sort((a, b) => a.shortName.localeCompare(b.shortName, 'zh-CN'));
}

function getCourseKeysForMajor(majorId) {
  return Array.from(new Set((state.timetable?.eventsByMajor[majorId] ?? []).map(getCourseKey)));
}

function getCourseKey(event) {
  return `${getMajorGroup(event.majorId)}::${normalizeCourseName(event.shortName || event.title)}`;
}

function getMajorGroup(majorId) {
  const text = normalize(majorId).toLowerCase();
  return text === '天河班' || text === 'th' ? 'tianhe' : 'regular';
}

function getMajorGroupLabel(majorId) {
  return getMajorGroup(majorId) === 'tianhe' ? '天河班' : '普通专业';
}

function getSelectedConflicts() {
  const selectedEvents = dedupeEventsByCourseAndTime(getAllEvents().filter((event) => state.selectedCourseKeys.has(getCourseKey(event))));
  const cells = new Map();
  for (const event of applyOccurrenceEdits(selectedEvents, false)) {
    const key = `${event.week}|${event.day}|${event.slot}`;
    if (!cells.has(key)) cells.set(key, new Map());
    cells.get(key).set(getCourseKey(event), event);
  }
  return Array.from(cells.entries()).map(([key, courses]) => {
    const [week, day, slot] = key.split('|');
    return { key, week, day, slot, courses: Array.from(courses.values()) };
  }).filter((item) => item.courses.length > 1).sort((a, b) => Number(a.week) - Number(b.week) || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || PERIODS.indexOf(a.slot) - PERIODS.indexOf(b.slot));
}

function getConflictCourseKeys(conflicts) {
  return new Set(conflicts.flatMap((conflict) => conflict.courses.map(getCourseKey)));
}

function dedupeEventsByCourseAndTime(events) {
  const byKey = new Map();
  for (const event of events) {
    const key = `${getCourseKey(event)}|${event.week}|${event.day}|${event.slot}`;
    const existing = byKey.get(key);
    if (!existing) {
      byKey.set(key, { ...event, sourceMajors: new Set([event.majorId]) });
      continue;
    }
    existing.sourceMajors.add(event.majorId);
    existing.teacher = mergeText(existing.teacher, event.teacher);
    existing.room = mergeText(existing.room, event.room);
  }
  return Array.from(byKey.values()).map((event) => ({ ...event, sourceMajors: Array.from(event.sourceMajors) }));
}

function applyOccurrenceEdits(events, filterWeek) {
  const edits = normalizeOccurrenceEdits(state.occurrenceEdits);
  const deleted = new Set(edits.deleted);
  const base = events.filter((event) => !deleted.has(getOccurrenceKey(event))).map((event) => {
    const key = getOccurrenceKey(event);
    return edits.roomOverrides[key] == null ? event : { ...event, room: edits.roomOverrides[key] };
  });
  const added = edits.added
    .filter((event) => !deleted.has(getOccurrenceKey(event)))
    .filter((event) => !filterWeek || !state.selectedWeek || event.week === state.selectedWeek)
    .map((event) => {
      const key = getOccurrenceKey(event);
      return edits.roomOverrides[key] == null ? event : { ...event, room: edits.roomOverrides[key] };
    });
  return [...base, ...added].sort(compareEvents);
}

function renderCoursePicker(catalog, conflicts) {
  const conflictCourseKeys = getConflictCourseKeys(conflicts);
  return `
    <section class="coursePicker">
      <div class="pickerHeader">
        <div><p class="eyebrow">课程选择</p><h2>按最终勾选课程生成课表</h2></div>
        <input id="courseSearch" class="searchInput" type="search" value="${escapeHtml(state.courseSearch)}" placeholder="搜索课程、教师、地点" />
      </div>
      <div class="courseList">
        ${catalog.map((course) => `
          <label class="courseOption ${state.selectedCourseKeys.has(course.key) ? 'selected' : ''} ${conflictCourseKeys.has(course.key) ? 'conflict' : ''}" data-course-search="${escapeHtml(`${course.majorLabels.join(' ')} ${course.title} ${course.shortName} ${course.teacher} ${course.room}`.toLowerCase())}">
            <input type="checkbox" data-course-key="${escapeHtml(course.key)}" ${state.selectedCourseKeys.has(course.key) ? 'checked' : ''} />
            <span>
              <strong>${escapeHtml(course.shortName)}</strong>
              <small>${escapeHtml(course.title)}</small>
              <small>${escapeHtml(course.groupLabel)} · ${escapeHtml(course.majorLabels.join(' / '))} · ${course.weekCount} 周</small>
              ${course.room ? `<small>${escapeHtml(course.room)}</small>` : ''}
              ${conflictCourseKeys.has(course.key) ? '<em>存在时间冲突</em>' : ''}
            </span>
          </label>
        `).join('')}
      </div>
    </section>
  `;
}

function filterCourseOptions(query) {
  const normalized = normalize(query).toLowerCase();
  document.querySelectorAll('[data-course-search]').forEach((option) => {
    option.hidden = normalized && !option.dataset.courseSearch.includes(normalized);
  });
}

function renderConflictSummary(conflicts) {
  const visible = conflicts.slice(0, 8);
  return `
    <section class="conflictPanel">
      <div><p class="eyebrow">冲突提示</p><h2>存在 ${conflicts.length} 个时间冲突</h2></div>
      <ul>${visible.map((conflict) => `<li><strong>第 ${escapeHtml(conflict.week)} 周 ${escapeHtml(conflict.day)} ${escapeHtml(conflict.slot)}</strong><span>${conflict.courses.map((event) => escapeHtml(event.shortName || event.title)).join(' / ')}</span></li>`).join('')}</ul>
    </section>
  `;
}

function renderSchedule(major, events, today) {
  const days = state.showWeekend ? DAYS : DAYS.slice(0, 5);
  const dateRange = getDateRangeForWeek(state.selectedWeek) || events[0]?.dateRange || '';
  const weekDates = getWeekDates(state.selectedWeek);
  const eventsByCell = new Map();
  for (const event of events) {
    const key = `${event.day}|${event.slot}`;
    if (!eventsByCell.has(key)) eventsByCell.set(key, []);
    eventsByCell.get(key).push(event);
  }

  return `
    <section class="scheduleHeader">
      <div><p class="eyebrow">${escapeHtml(major?.label ?? '')}</p><h2>第 ${escapeHtml(state.selectedWeek || '-')} 周</h2></div>
      <p>${escapeHtml(dateRange)}</p>
    </section>
    <section class="wakeupBoard ${state.selectedWeek === today?.week ? 'currentWeek' : ''}" style="--day-count:${days.length}" data-swipe-week>
      <div class="boardCorner">${escapeHtml(getMonthLabel(dateRange))}</div>
      ${days.map((day, index) => `
        <div class="boardDay ${state.selectedWeek === today?.week && day === today.day ? 'todayColumn' : ''}" style="grid-column:${index + 2};">
          <strong>${escapeHtml(day.replace('星期', ''))}</strong>
          <span>${escapeHtml(weekDates[index] ?? '')}</span>
        </div>
      `).join('')}
      ${PERIOD_DETAILS.map((period) => `
        <div class="timeHead" style="grid-row:${period.index + 1};">
          <strong>${period.index}</strong>
          <span>${period.start}</span>
          <span>${period.end}</span>
        </div>
      `).join('')}
      ${days.flatMap((day, dayIndex) => PERIODS.map((slot) => renderEmptyBlock(day, dayIndex, slot, today))).join('')}
      ${days.flatMap((day, dayIndex) => PERIODS.map((slot) => renderCourseBlock(eventsByCell.get(`${day}|${slot}`) ?? [], day, dayIndex, slot, today))).join('')}
    </section>
  `;
}

function renderEmptyBlock(day, dayIndex, slot, today) {
  const meta = getSlotMeta(slot);
  if (!meta) return '';
  return `<button class="emptyBlock" type="button" data-open-cell="${escapeHtml(`${day}|${slot}`)}" aria-label="添加课程" style="grid-column:${dayIndex + 2};grid-row:${meta.start + 1} / span ${meta.span};"></button>`;
}

function renderCourseBlock(events, day, dayIndex, slot, today) {
  const meta = getSlotMeta(slot);
  if (!meta || !events.length) return '';
  const cellKey = `${day}|${slot}`;
  const isToday = state.selectedWeek === today?.week && day === today.day;
  if (events.length > 1) {
    return `<button class="courseBlock conflictBlock ${isToday ? 'todayCell' : ''}" type="button" data-open-conflict="${escapeHtml(cellKey)}" style="grid-column:${dayIndex + 2};grid-row:${meta.start + 1} / span ${meta.span};"><strong>出现冲突</strong><span>${events.length} 门课程</span><small>${events.map((event) => escapeHtml(event.shortName || event.title)).join(' / ')}</small></button>`;
  }
  const event = events[0];
  return `<button class="courseBlock ${getCourseColorClass(event)} ${isToday ? 'todayCell' : ''}" type="button" data-open-cell="${escapeHtml(cellKey)}" style="grid-column:${dayIndex + 2};grid-row:${meta.start + 1} / span ${meta.span};"><strong>${escapeHtml(event.shortName || event.title)}</strong>${event.room ? `<span>@${escapeHtml(event.room)}</span>` : ''}${event.teacher ? `<small>${escapeHtml(event.teacher)}</small>` : ''}</button>`;
}

function renderConflictModal(events) {
  if (!state.activeConflictKey) return '';
  const conflictEvents = events.filter((event) => `${event.day}|${event.slot}` === state.activeConflictKey);
  if (conflictEvents.length <= 1) return '';
  const [day, slot] = state.activeConflictKey.split('|');
  return `
    <div class="modalBackdrop" role="presentation">
      <section class="conflictModal" role="dialog" aria-modal="true">
        <div class="modalHeader"><div><p class="eyebrow">冲突详情</p><h2>第 ${escapeHtml(state.selectedWeek || '-')} 周 ${escapeHtml(day)} ${escapeHtml(slot)}</h2></div><button class="modalClose" type="button" data-close-modal>关闭</button></div>
        <div class="modalCourseList">${conflictEvents.map(renderModalCourse).join('')}</div>
      </section>
    </div>
  `;
}

function renderCellEditModal(events, catalog) {
  if (!state.activeEditCellKey) return '';
  const [day, slot] = state.activeEditCellKey.split('|');
  const cellEvents = events.filter((event) => `${event.day}|${event.slot}` === state.activeEditCellKey);
  return `
    <div class="modalBackdrop" role="presentation">
      <section class="conflictModal editModal" role="dialog" aria-modal="true">
        <div class="modalHeader"><div><p class="eyebrow">单次调整</p><h2>第 ${escapeHtml(state.selectedWeek || '-')} 周 ${escapeHtml(day)} ${escapeHtml(slot)}</h2></div><button class="modalClose" type="button" data-close-modal>关闭</button></div>
        ${cellEvents.length ? renderExistingOccurrenceEditor(cellEvents[0]) : renderAddOccurrenceEditor(catalog)}
      </section>
    </div>
  `;
}

function renderModalCourse(event) {
  return `<article class="modalCourse"><strong>${escapeHtml(event.shortName || event.title)}</strong><span>${escapeHtml(event.title)}</span>${event.room ? `<small>地点：${escapeHtml(event.room)}</small>` : ''}${event.teacher ? `<small>教师：${escapeHtml(event.teacher)}</small>` : ''}${event.hours ? `<small>学时：${escapeHtml(event.hours)}</small>` : ''}</article>`;
}

function renderExistingOccurrenceEditor(event) {
  return `${renderModalCourse(event)}<form id="editRoomForm" class="editForm"><input type="hidden" name="occurrenceKey" value="${escapeHtml(getOccurrenceKey(event))}" /><label><span>修改这一次课的教室</span><input class="textInput" name="room" type="text" value="${escapeHtml(event.room)}" placeholder="填写新教室" /></label><div class="modalActions"><button class="primaryButton" type="submit">保存教室</button><button id="deleteOccurrenceButton" class="ghostButton danger" type="button" data-occurrence-key="${escapeHtml(getOccurrenceKey(event))}">删除这一次课</button></div></form>`;
}

function renderAddOccurrenceEditor(catalog) {
  return `
    <form id="addExistingCourseForm" class="editForm"><label><span>添加已有课程</span><select name="courseKey">${catalog.map((course) => `<option value="${escapeHtml(course.key)}">${escapeHtml(course.groupLabel)} · ${escapeHtml(course.shortName)}</option>`).join('')}</select></label><label><span>教室</span><input class="textInput" name="room" type="text" placeholder="可留空，默认使用原教室" /></label><button class="primaryButton" type="submit">添加已有课程</button></form>
    <form id="addCustomCourseForm" class="editForm customCourseForm"><p class="fieldLabel">自定义课程</p><label><span>课程名称</span><input class="textInput" name="title" type="text" required placeholder="必填" /></label><label><span>简称</span><input class="textInput" name="shortName" type="text" /></label><label><span>教室</span><input class="textInput" name="room" type="text" required placeholder="必填" /></label><label><span>教师</span><input class="textInput" name="teacher" type="text" /></label><button class="primaryButton" type="submit">添加自定义课程</button></form>
  `;
}

async function onEditRoomSubmit(event) {
  event.preventDefault();
  const data = new FormData(event.currentTarget);
  const key = normalize(data.get('occurrenceKey'));
  const room = normalize(data.get('room'));
  state.occurrenceEdits = normalizeOccurrenceEdits(state.occurrenceEdits);
  if (room) state.occurrenceEdits.roomOverrides[key] = room;
  else delete state.occurrenceEdits.roomOverrides[key];
  await saveAppState();
  state.activeEditCellKey = '';
  render();
}

async function onDeleteOccurrence(event) {
  const key = event.currentTarget.dataset.occurrenceKey;
  state.occurrenceEdits = normalizeOccurrenceEdits(state.occurrenceEdits);
  if (!state.occurrenceEdits.deleted.includes(key)) state.occurrenceEdits.deleted.push(key);
  delete state.occurrenceEdits.roomOverrides[key];
  await saveAppState();
  state.activeEditCellKey = '';
  render();
}

async function onAddExistingCourse(event) {
  event.preventDefault();
  const data = new FormData(event.currentTarget);
  const course = getCourseCatalog().find((item) => item.key === normalize(data.get('courseKey')));
  if (!course) return;
  const [day, slot] = state.activeEditCellKey.split('|');
  await addManualEvent({
    majorId: course.group === 'tianhe' ? '天河班' : state.selectedMajor || '自定义',
    week: state.selectedWeek,
    dateRange: getDateRangeForWeek(state.selectedWeek),
    day,
    slot,
    title: course.title,
    shortName: course.shortName,
    teacher: course.teacher || '',
    room: normalize(data.get('room')) || course.room || '',
    hours: ''
  });
}

async function onAddCustomCourse(event) {
  event.preventDefault();
  const data = new FormData(event.currentTarget);
  const [day, slot] = state.activeEditCellKey.split('|');
  const title = normalize(data.get('title'));
  const room = normalize(data.get('room'));
  if (!title || !room) return;
  await addManualEvent({
    majorId: state.selectedMajor || '自定义',
    week: state.selectedWeek,
    dateRange: getDateRangeForWeek(state.selectedWeek),
    day,
    slot,
    title,
    shortName: normalize(data.get('shortName')) || title,
    teacher: normalize(data.get('teacher')),
    room,
    hours: ''
  });
}

async function addManualEvent(event) {
  state.occurrenceEdits = normalizeOccurrenceEdits(state.occurrenceEdits);
  const manualEvent = { ...event, id: `manual-${createId()}` };
  state.occurrenceEdits.deleted = state.occurrenceEdits.deleted.filter((item) => item !== getOccurrenceKey(manualEvent));
  state.occurrenceEdits.added.push(manualEvent);
  await saveAppState();
  state.activeEditCellKey = '';
  render();
}

function renderCsvNotice() {
  return `
    <section class="csvNotice">
      <div><p class="eyebrow">CSV 标准格式</p><h2>按模板导入自定义课程</h2></div>
      <div class="csvNoticeBody"><p>必填：<code>专业,周次,星期,节次,课程名称</code>。推荐补充：<code>日期,简称,教师,教室,学时</code>。</p><div class="csvHint"><strong>示例</strong><code>人工智能,1,03.09-03.15,星期一,1/2,机器学习,机器学习,张老师,302-106,80</code></div></div>
      <a class="downloadLink" href="/csv-template.csv" download="课程表CSV模板.csv">下载模板</a>
    </section>
  `;
}

function getCurrentAcademicDay() {
  if (!state.timetable) return null;
  const today = new Date();
  const year = state.timetable.termYear || today.getFullYear();
  for (const week of getAvailableWeeks()) {
    const range = getDateRangeForWeek(week);
    const parsed = parseDateRange(range, year);
    if (!parsed) continue;
    const start = startOfDay(parsed.start);
    const end = startOfDay(parsed.end);
    const now = startOfDay(today);
    if (now >= start && now <= end) {
      const dayIndex = Math.min(Math.max(Math.round((now - start) / 86400000), 0), 6);
      return { week, day: DAYS[dayIndex] };
    }
  }
  return null;
}

function getWeekDates(week) {
  const range = getDateRangeForWeek(week);
  const parsed = parseDateRange(range, state.timetable?.termYear || new Date().getFullYear());
  if (!parsed) return [];
  return DAYS.map((_, index) => {
    const date = new Date(parsed.start);
    date.setDate(parsed.start.getDate() + index);
    return String(date.getDate());
  });
}

function getMonthLabel(range) {
  const match = normalize(range).match(/^(\d{1,2})/);
  return match ? `${Number(match[1])}月` : '';
}

function getSlotMeta(slot) {
  const match = normalize(slot).match(/(\d+)[/-](\d+)/);
  if (!match) return null;
  const start = Number(match[1]);
  const end = Number(match[2]);
  return { start, span: Math.max(end - start + 1, 1) };
}

function getCourseColorClass(event) {
  const key = normalizeCourseName(event.shortName || event.title);
  let hash = 0;
  for (const char of key) hash = (hash * 31 + char.charCodeAt(0)) % 997;
  return `courseTone${(hash % 6) + 1}`;
}

function getDateRangeForWeek(week) {
  const event = getAllEvents().find((item) => item.week === week && item.dateRange);
  return event?.dateRange ?? '';
}

function parseDateRange(value, year) {
  const match = normalize(value).match(/(\d{1,2})[.月/-](\d{1,2}).*?(\d{1,2})[.月/-](\d{1,2})/);
  if (!match) return null;
  const startMonth = Number(match[1]);
  const startDay = Number(match[2]);
  let endMonth = Number(match[3]);
  const endDay = Number(match[4]);
  let endYear = year;
  if (endMonth < startMonth) endYear += 1;
  return { start: new Date(year, startMonth - 1, startDay), end: new Date(endYear, endMonth - 1, endDay) };
}

function inferTermYear(text) {
  const full = normalize(text).match(/20\d{2}/);
  if (full) return Number(full[0]);
  const short = normalize(text).match(/(?:^|[^\d])(\d{2})\s*(?:春|秋|学期)/);
  if (short) return 2000 + Number(short[1]);
  return new Date().getFullYear();
}

function compareEvents(a, b) {
  return Number(a.week) - Number(b.week) || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || PERIODS.indexOf(a.slot) - PERIODS.indexOf(b.slot);
}

function getOccurrenceKey(event) {
  return `${getCourseKey(event)}|${event.week}|${event.day}|${event.slot}`;
}

function normalizeOccurrenceEdits(edits = {}) {
  return {
    deleted: Array.isArray(edits.deleted) ? edits.deleted : [],
    roomOverrides: edits.roomOverrides && typeof edits.roomOverrides === 'object' ? edits.roomOverrides : {},
    added: Array.isArray(edits.added) ? edits.added : []
  };
}

function cloneOccurrenceEdits(edits) {
  const normalized = normalizeOccurrenceEdits(edits);
  return { deleted: [...normalized.deleted], roomOverrides: { ...normalized.roomOverrides }, added: normalized.added.map((event) => ({ ...event })) };
}

function persistCurrentSchedule() {
  if (!state.timetable) return;
  const schedule = serializeCurrentSchedule();
  state.activeScheduleId = schedule.id;
  const index = state.schedules.findIndex((item) => item.id === schedule.id);
  if (index >= 0) state.schedules[index] = schedule;
  else state.schedules.push(schedule);
}

async function persistCurrentScheduleDraft() {
  if (!state.timetable) return;
  await saveAppState();
}

function serializeCurrentSchedule() {
  return {
    id: state.activeScheduleId || createId(),
    name: state.scheduleName || state.sourceName || '未命名课表',
    sourceName: state.sourceName,
    timetable: state.timetable,
    selectedMajor: state.selectedMajor,
    selectedWeek: state.selectedWeek,
    selectedCourseKeys: Array.from(state.selectedCourseKeys),
    showWeekend: state.showWeekend,
    occurrenceEdits: cloneOccurrenceEdits(state.occurrenceEdits),
    updatedAt: new Date().toISOString()
  };
}

function applySchedule(schedule) {
  state.activeScheduleId = schedule.id;
  state.scheduleName = schedule.name ?? schedule.sourceName ?? '未命名课表';
  state.sourceName = schedule.sourceName ?? '';
  state.timetable = schedule.timetable;
  state.selectedMajor = schedule.selectedMajor ?? '';
  state.selectedCourseKeys = migrateSavedCourseKeys(schedule.selectedCourseKeys ?? []);
  state.showWeekend = schedule.showWeekend !== false;
  state.occurrenceEdits = normalizeOccurrenceEdits(schedule.occurrenceEdits);
  state.selectedWeek = getCurrentAcademicDay()?.week ?? schedule.selectedWeek ?? '';
  state.setupMode = false;
  state.courseSearch = '';
  state.activeConflictKey = '';
  state.activeEditCellKey = '';
}

function resetCurrentSchedule() {
  state.activeScheduleId = '';
  state.scheduleName = '';
  state.sourceName = '';
  state.timetable = null;
  state.selectedMajor = '';
  state.selectedWeek = '';
  state.selectedCourseKeys = new Set();
  state.occurrenceEdits = normalizeOccurrenceEdits();
  state.setupMode = false;
  state.saveStatus = '';
  state.courseSearch = '';
  state.activeConflictKey = '';
  state.activeEditCellKey = '';
}

function normalizeSavedSchedules(savedState) {
  if (Array.isArray(savedState.schedules)) {
    return savedState.schedules.map((schedule) => ({ ...schedule, id: schedule.id || createId(), name: schedule.name || schedule.sourceName || '未命名课表' }));
  }
  if (savedState.timetable) {
    return [{
      id: createId(),
      name: savedState.sourceName || '已保存课表',
      sourceName: savedState.sourceName ?? '',
      timetable: savedState.timetable,
      selectedMajor: savedState.selectedMajor ?? '',
      selectedWeek: savedState.selectedWeek ?? '',
      selectedCourseKeys: savedState.selectedCourseKeys ?? [],
      showWeekend: savedState.showWeekend !== false,
      occurrenceEdits: normalizeOccurrenceEdits(savedState.occurrenceEdits),
      updatedAt: savedState.savedAt ?? new Date().toISOString()
    }];
  }
  return [];
}

function migrateSavedCourseKeys(keys) {
  return new Set(keys.flatMap((key) => {
    const text = normalize(key);
    if (text.startsWith('regular::') || text.startsWith('tianhe::')) return [text];
    if (text.includes('::')) {
      const [legacyMajor, legacyName] = text.split('::');
      return [`${getMajorGroup(legacyMajor)}::${normalizeCourseName(legacyName)}`];
    }
    return [`${getMajorGroup(state.selectedMajor)}::${normalizeCourseName(text)}`];
  }).filter(Boolean));
}

async function saveAppState() {
  persistCurrentSchedule();
  const db = await openDb();
  await requestToPromise(db.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).put({
    version: 3,
    activeScheduleId: state.activeScheduleId,
    schedules: state.schedules,
    savedAt: new Date().toISOString()
  }, APP_STATE_KEY));
  db.close();
}

async function loadSavedState() {
  try {
    const db = await openDb();
    const result = await requestToPromise(db.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).get(APP_STATE_KEY));
    db.close();
    return result;
  } catch {
    return null;
  }
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function pick(row, keys) {
  for (const key of keys) {
    const value = row[key] ?? row[normalize(key)];
    if (normalize(value)) return normalize(value);
  }
  return '';
}

function normalize(value) {
  return value == null ? '' : String(value).trim();
}

function normalizeCourseName(value) {
  return normalize(value).replace(/\s+/g, '').toLowerCase();
}

function normalizeDay(value) {
  const text = normalize(value);
  const map = {
    '1': '星期一', 一: '星期一', 周一: '星期一', 星期一: '星期一',
    '2': '星期二', 二: '星期二', 周二: '星期二', 星期二: '星期二',
    '3': '星期三', 三: '星期三', 周三: '星期三', 星期三: '星期三',
    '4': '星期四', 四: '星期四', 周四: '星期四', 星期四: '星期四',
    '5': '星期五', 五: '星期五', 周五: '星期五', 星期五: '星期五',
    '6': '星期六', 六: '星期六', 周六: '星期六', 星期六: '星期六',
    '7': '星期日', 日: '星期日', 天: '星期日', 周日: '星期日', 周天: '星期日', 星期日: '星期日', 星期天: '星期日', 星期七: '星期日'
  };
  return map[text] ?? text;
}

function normalizeSlot(value) {
  const text = normalize(value);
  if (PERIODS.includes(text)) return text;
  const match = text.replace(/\s/g, '').match(/(\d+)[/-](\d+)/);
  return match ? `${match[1]}/${match[2]}` : text;
}

function mergeText(left, right) {
  return Array.from(new Set([left, right].flatMap((item) => normalize(item).split(',')).map(normalize).filter(Boolean))).join(',');
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDate(date) {
  return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

function createDefaultScheduleName(fileName) {
  return normalize(fileName).replace(/\.(xlsx|xls|csv)$/i, '') || '新课表';
}

function createId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function escapeHtml(value) {
  return normalize(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('/sw.js').catch(() => {});
    });
  }
}

const $ = id => document.getElementById(id);
let currentPage = 1;
const pageSize = 50;
let searchTimer = null;

async function api(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok) throw Error(data.error || '请求失败');
  return data;
}

async function adminLogin() {
  try {
    await api('/api/auth/login', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({username: $('adminUser').value, password: $('adminPass').value})
    });
    await checkAuth();
  } catch (error) {
    $('loginMsg').textContent = error.message;
  }
}

async function adminLogout() {
  await api('/api/auth/logout', {method: 'POST'});
  location.reload();
}

function showPage(page) {
  document.querySelectorAll('.admin-page').forEach(section => {
    section.hidden = section.dataset.page !== page;
  });
  document.querySelectorAll('.admin-tabs .tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.page === page);
  });
  if (page === 'questions') load(1);
}

function queueLoad() {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => load(1), 250);
}

async function checkAuth() {
  const user = await api('/api/auth/me');
  if (user.role !== 'admin') {
    $('adminLogin').hidden = false;
    return;
  }
  $('adminLogin').hidden = true;
  $('adminContent').hidden = false;
  await loadStats();
  await loadStudentOverview();
  await loadUsers();
}

async function loadStudentOverview() {
  const stats = await api('/api/admin/overview-stats');
  $('studentOverview').textContent = `学生数：${stats.students}　答题人次：${stats.attempts}　总答题题数：${stats.total_questions}　平均正确率：${stats.average_accuracy}%`;
}
async function loadUsers() {
  const users = await api('/api/admin/users');
  const root = $('usersList'); root.replaceChildren(); root.append(createEl('h3', '学生账户管理'));
  users.filter(user => user.role === 'student').forEach(user => {
    const row = createEl('div'); row.className = 'detail'; row.append(createEl('span', `${user.username} `));
    const input = document.createElement('input'); input.type = 'password'; input.placeholder = '新密码（至少6位）';
    const button = createEl('button', '重置密码'); button.addEventListener('click', async () => { try { await api(`/api/admin/users/${user.id}/reset-password`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({password:input.value})}); input.value=''; alert('密码已重置'); } catch (error) { alert(error.message); } });
    row.append(input, button); root.append(row);
  });
}
async function loadStats() {
  const stats = await api('/api/stats');
  $('stats').innerHTML = `题目总数：<b>${Number(stats.total) || 0}</b>　` +
    Object.entries(stats.types).map(item => `${escapeHtml(item[0])}：${Number(item[1]) || 0}`).join('　') +
    '<br><small>' + stats.chapters.map(item => `${escapeHtml(item.chapter)}（${item.count}）`).join('　') + '</small>';

  const currentChapter = $('chapterFilter').value;
  $('chapterFilter').innerHTML = '<option value="">全部章节</option>' +
    stats.chapters.map(item => `<option value="${escapeHtml(item.chapter)}">${escapeHtml(item.chapter)}（${item.count}）</option>`).join('');
  $('chapterFilter').value = currentChapter;
}

async function changePassword() {
  const current = $('currentPassword').value;
  const newPassword = $('newPassword').value;
  try {
    await api('/api/auth/change-password', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({current_password: current, new_password: newPassword})
    });
    $('passwordMsg').textContent = '密码已修改，请重新登录';
    setTimeout(() => location.reload(), 700);
  } catch (error) {
    $('passwordMsg').textContent = error.message;
  }
}

function selectedTypes() {
  return [...document.querySelectorAll('.filter-types input:checked')].map(input => input.value);
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
  }[character]));
}

function renderQuestions(items) {
  const list = $('list');
  list.replaceChildren();
  if (!items.length) {
    list.innerHTML = '<div class="empty-state">没有找到符合条件的题目</div>';
    return;
  }
  $('selectAllQuestions').checked = false;
  items.forEach(question => {
    const detail = document.createElement('div');
    detail.className = 'detail';
    const checkbox = document.createElement('input'); checkbox.type = 'checkbox'; checkbox.className = 'question-select'; checkbox.value = question.id;
    detail.append(checkbox, createEl('b', `#${question.id} ${question.type}　${question.chapter}`), createEl('div', question.question), createEl('div', `答案：${question.answer}`));

    const editButton = document.createElement('button');
    editButton.textContent = '编辑';
    editButton.addEventListener('click', () => edit(question));
    const deleteButton = document.createElement('button');
    deleteButton.className = 'secondary';
    deleteButton.textContent = '删除';
    deleteButton.addEventListener('click', () => del(question.id));
    detail.append(editButton, deleteButton);
    list.append(detail);
  });
}

async function load(page = 1) {
  try {
    currentPage = Math.max(1, page);
    const params = new URLSearchParams({
      search: $('search').value,
      chapter: $('chapterFilter').value,
      page: currentPage,
      page_size: pageSize
    });
    selectedTypes().forEach(type => params.append('types', type));

    const result = await api('/api/questions?' + params);
    currentPage = result.page;
    renderQuestions(result.items || []);

    const totalPages = Math.max(1, Math.ceil(result.total / result.page_size));
    $('pageInfo').textContent = `第 ${result.page} / ${totalPages} 页`;
    $('pageSummary').textContent = `共 ${result.total} 题，每页 ${result.page_size} 题`;
    $('prevPage').disabled = result.page <= 1;
    $('nextPage').disabled = result.page >= totalPages;
  } catch (error) {
    $('list').textContent = error.message;
  }
}

function toggleAllQuestions(checked) { document.querySelectorAll('.question-select').forEach(input => { input.checked = checked; }); }
function selectedQuestionIds() { return [...document.querySelectorAll('.question-select:checked')].map(input => Number(input.value)); }
async function batchDelete() {
  const ids = selectedQuestionIds();
  if (!ids.length || !confirm(`确认删除 ${ids.length} 道题目？`)) return;
  try {
    await api('/api/questions/batch-delete', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ids})});
    await loadStats();
    await load(currentPage);
  } catch (error) {
    $('msg').textContent = error.message;
  }
}
async function batchUpdate() {
  const ids = selectedQuestionIds();
  const chapter = $('batchChapter').value.trim();
  if (!ids.length || !chapter) return alert('请选择题目并填写章节');
  try {
    await api('/api/questions/batch-update', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ids, chapter})});
    $('batchChapter').value = '';
    await loadStats();
    await load(currentPage);
  } catch (error) {
    $('msg').textContent = error.message;
  }
}
async function restoreBackup() { const file=$('backupFile').files[0]; if (!file || !confirm('恢复前会自动备份当前题库，确认恢复？')) return; try { const data=JSON.parse(await file.text()); const result=await api('/api/questions/restore',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}); $('msg').textContent=`恢复 ${result.restored} 道题目完成`; await loadStats(); await load(1); } catch(error) { $('msg').textContent=error.message; } }
async function loadImportHistory() { const rows=await api('/api/import-history'); $('importHistory').textContent=rows.map(row=>`${row.created_at} ${row.source} 导入${row.imported} 跳过${row.skipped} 重复${row.duplicates}`).join('\n') || '暂无导入记录'; }

async function clearQuestionBank() {
  const confirmed = confirm('确定要清空全部题目吗？此操作不可撤销，并会同时删除学生的错题和收藏关联数据。');
  if (!confirmed) return;
  try {
    const result = await api('/api/questions/clear', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'}
    });
    $('msg').textContent = `题库已清空，共删除 ${result.deleted} 道题目`;
    await loadStats();
    await load(1);
  } catch (error) {
    $('msg').textContent = error.message;
  }
}

async function importData(clear) {
  try {
    const result = await api('/api/import', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({clear})
    });
    $('msg').textContent = `已处理 ${result.files} 个文件，导入 ${result.imported} 题，跳过 ${result.skipped} 行，重复 ${result.duplicates || 0} 题`;
    await loadStats();
    await load(1);
  } catch (error) {
    $('msg').textContent = error.message;
  }
}

const fieldAliases = {
  question: ['question', '题干', '题目', '问题'],
  answer: ['answer', '答案', '正确答案'],
  A: ['a', '选项a', 'a选项'], B: ['b', '选项b', 'b选项'],
  C: ['c', '选项c', 'c选项'], D: ['d', '选项d', 'd选项'], E: ['e', '选项e', 'e选项'],
  chapter: ['chapter', '章节', '单元'], explanation: ['explanation', '解析', '答案解析']
};
let excelSheets = [];

function normalizeHeader(value) { return String(value ?? '').trim().toLowerCase().replace(/[\s：:（）()]/g, ''); }
function createEl(tag, text) { const element = document.createElement(tag); if (text !== undefined) element.textContent = text; return element; }
function addFieldSelect(parent, label, headers, selected, field) {
  const wrap = createEl('label', label + ' '), select = document.createElement('select');
  select.dataset.field = field;
  const none = createEl('option', '不映射'); none.value = ''; select.append(none);
  headers.forEach((header, index) => { const option = createEl('option', String(header ?? `列${index + 1}`)); option.value = String(index); if (String(index) === String(selected)) option.selected = true; select.append(option); });
  wrap.append(select); parent.append(wrap);
}
function detectColumn(headers, field) { return headers.findIndex(header => fieldAliases[field].map(normalizeHeader).includes(normalizeHeader(header))); }

async function previewExcel() {
  excelSheets = []; $('excelPreview').replaceChildren(); $('confirmExcelImport').hidden = true; $('excelMsg').textContent = '';
  const files = [...$('uploadFiles').files]; if (!files.length) return;
  if (!window.XLSX) { $('excelMsg').textContent = 'Excel解析组件未加载'; return; }
  try {
    for (const file of files) {
      const workbook = XLSX.read(await file.arrayBuffer(), {type: 'array'});
      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName], rows = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: '', raw: false});
        if (!rows.length) continue;
        let headers = rows[0].map(value => String(value ?? '')), start = 1;
        if (!headers.some(header => detectColumn(headers, 'question') >= 0)) { headers = Object.keys(rows[0] || {}); start = 0; }
        excelSheets.push({file, sheetName, headers, rows: rows.slice(start), type: /多选|mchoice/i.test(sheetName) ? '多选' : /判断|judgement/i.test(sheetName) ? '判断' : '单选', map: Object.fromEntries(['question','answer','A','B','C','D','E','chapter','explanation'].map(field => [field, detectColumn(headers, field)]))});
      }
    }
    renderExcelPreview(); $('confirmExcelImport').hidden = excelSheets.length === 0; $('excelMsg').textContent = `共识别 ${excelSheets.length} 个工作表`;
  } catch (error) { $('excelMsg').textContent = '读取Excel失败：' + error.message; }
}

function renderExcelPreview() {
  const root = $('excelPreview'); root.replaceChildren();
  excelSheets.forEach((item, index) => {
    const section = createEl('section'); section.className = 'detail'; section.append(createEl('h4', `${item.file.name} / ${item.sheetName}（${item.rows.length} 行）`));
    const typeLabel = createEl('label', '题型 '), type = document.createElement('select'); type.dataset.sheet = index;
    ['单选', '多选', '判断'].forEach(value => { const option = createEl('option', value); option.value = value; option.selected = value === item.type; type.append(option); });
    type.addEventListener('change', () => item.type = type.value); typeLabel.append(type); section.append(typeLabel);
    const mapping = createEl('div'); ['question','answer','A','B','C','D','E','chapter','explanation'].forEach(field => addFieldSelect(mapping, {question:'题干',answer:'答案',A:'A选项',B:'B选项',C:'C选项',D:'D选项',E:'E选项',chapter:'章节',explanation:'解析'}[field], item.headers, item.map[field], field));
    mapping.addEventListener('change', event => { if (event.target.matches('select[data-field]')) item.map[event.target.dataset.field] = event.target.value === '' ? -1 : Number(event.target.value); }); section.append(mapping);
    const table = document.createElement('table'), head = document.createElement('tr'); item.headers.forEach(header => head.append(createEl('th', String(header)))); table.append(head);
    item.rows.slice(0, 5).forEach(row => { const tr = document.createElement('tr'); item.headers.forEach((_, column) => tr.append(createEl('td', String(row[column] ?? '')))); table.append(tr); }); section.append(table); root.append(section);
  });
}

async function confirmExcelImport() {
  const questions = [];
  for (const item of excelSheets) {
    if (item.map.question < 0 || item.map.answer < 0) { $('excelMsg').textContent = `${item.sheetName}：请至少映射题干和答案`; return; }
    item.rows.forEach(row => {
      const question = {type: item.type, question: String(row[item.map.question] ?? '').trim(), answer: String(row[item.map.answer] ?? '').trim(), chapter: item.map.chapter >= 0 ? String(row[item.map.chapter] ?? '').trim() : '', explanation: item.map.explanation >= 0 ? String(row[item.map.explanation] ?? '').trim() : '', options: {}};
      ['A','B','C','D','E'].forEach(key => { const column = item.map[key]; if (column >= 0 && String(row[column] ?? '').trim()) question.options[key] = String(row[column]).trim(); });
      if (question.question && question.answer) questions.push(question);
    });
  }
  if (!questions.length) { $('excelMsg').textContent = '没有可导入的有效题目'; return; }
  try {
    const result = await api('/api/import-mapped', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({questions})});
    $('excelMsg').textContent = `导入完成：${result.imported} 题，跳过 ${result.skipped || 0} 行，重复 ${result.duplicates || 0} 题`; $('confirmExcelImport').hidden = true; await loadStats(); await load(1);
  } catch (error) { $('excelMsg').textContent = error.message; }
}

function edit(question = {}) {
  $('modal').classList.remove('hidden'); $('qid').value = question.id || ''; $('qtype').value = question.type || '单选'; $('chapter').value = question.chapter || ''; $('question').value = question.question || ''; $('answer').value = question.answer || '';
  const options = question.options || {}; ['A','B','C','D','E'].forEach(key => $(key).value = options[key] || ''); $('explanation').value = question.explanation || '';
}
function closeModal() { $('modal').classList.add('hidden'); }
async function save() {
  const options = {}; ['A','B','C','D','E'].forEach(key => { if ($(key).value) options[key] = $(key).value; });
  const data = {type: $('qtype').value, chapter: $('chapter').value, question: $('question').value, answer: $('answer').value, options, explanation: $('explanation').value};
  const id = $('qid').value;
  try { await api('/api/questions/' + (id || '0'), {method: id ? 'PUT' : 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)}); closeModal(); await loadStats(); await load(currentPage); }
  catch (error) { alert(error.message); }
}
async function del(id) { if (confirm('确认删除该题？')) { await api('/api/questions/' + id, {method: 'DELETE'}); await loadStats(); await load(currentPage); } }

$('uploadFiles').addEventListener('change', previewExcel);
checkAuth().catch(error => $('loginMsg').textContent = error.message);

import './style.css';
import type { Workbook, Worksheet, Cell, CellValue, Style } from 'exceljs';
import type ExcelJS from 'exceljs';

const icons: Record<string, string> = {
  upload: '<path d="M12 16V4m-4 4 4-4 4 4M4 16v4h16v-4"/>',
  file: '<path d="M14 2H5v20h14V7l-5-5Z M14 2v6h5 M8 12h8M8 16h8"/>',
  lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3m-4 4v3"/>',
  arrow: '<path d="M5 12h14m-5-5 5 5-5 5"/>',
  back: '<path d="m14 6-6 6 6 6"/>',
  download: '<path d="M12 3v12m-4-4 4 4 4-4M4 16v5h16v-5"/>',
  undo: '<path d="M4 10h10a6 6 0 0 1 0 12M4 10l5-5M4 10l5 5"/>',
  redo: '<path d="M20 10H10a6 6 0 0 0 0 12m10-12-5-5m5 5-5 5"/>',
  search: '<circle cx="10" cy="10" r="6"/><path d="m15 15 5 5"/>',
  grid: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>',
  feather: '<path d="M4 20 17 7M8 16c-1-8 7-15 12-12 3 5-4 13-12 12ZM8 12h4m0-4h4"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  help: '<circle cx="12" cy="12" r="9"/><path d="M9 8a3 3 0 0 1 6 0c0 2-3 2-3 5m0 3v1"/>',
};
const icon = (name: string) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] ?? ''}</svg>`;
const app = document.querySelector<HTMLDivElement>('#app')!;
let Excel: typeof ExcelJS;
let book: Workbook;
let sheet: Worksheet;
let filename = 'Sem título.xlsx';
let dirty = false;
let revision = 0;
let nextRevision = 0;
let savedRevision = 0;
let loading = false;
let row = 1, col = 1, anchorRow = 1, anchorCol = 1;
let rowCount = 100, colCount = 26;
let grid: HTMLDivElement;
let space: HTMLDivElement;
let editorInput: HTMLInputElement | null = null;
let pendingFormula: { r: number; c: number; text: string } | null = null;
let draggingSelection = false;
let lastPointer: { r: number; c: number; time: number } | null = null;
let toastTimer: ReturnType<typeof setTimeout>;
const ROW = 32, COL = 128, HEADER = 32, GUTTER = 48;
type Snapshot = { row: number; col: number; value: CellValue; style: Partial<Style> };
type Change = { sheetId: number; before: Snapshot[]; after: Snapshot[]; beforeRevision: number; afterRevision: number };
let undoStack: Change[] = [], redoStack: Change[] = [];

function escape(value: unknown) {
  return String(value ?? '').replace(/[&<>"']/g, x => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[x]!);
}
function toast(message: string) {
  document.querySelector('.toast')?.remove();
  const element = document.createElement('div');
  element.className = 'toast'; element.role = 'status'; element.textContent = message;
  document.body.append(element);
  clearTimeout(toastTimer); toastTimer = setTimeout(() => element.remove(), 5000);
}
async function busy<T>(message: string, task: () => Promise<T>): Promise<T | undefined> {
  if (loading) return;
  loading = true;
  const overlay = document.createElement('div');
  overlay.className = 'busy'; overlay.innerHTML = `<span class="spinner"></span>${escape(message)}`;
  overlay.role = 'status'; document.body.append(overlay);
  try { return await task(); }
  catch (error) { console.error(error); toast(error instanceof Error ? error.message : 'Não foi possível concluir. Tente novamente.'); }
  finally { loading = false; overlay.remove(); }
}
async function engine() {
  if (!Excel) Excel = (await import('exceljs')).default;
  return Excel;
}
function header() {
  return `<header class="app-header"><button class="brand" id="brand" aria-label="Folha, início"><span class="logo">f.</span>folha<small>PLANILHAS, SIMPLES ASSIM</small></button><div class="header-actions"><span class="privacy">${icon('lock')} Seus arquivos ficam com você</span><button class="help-button" id="help">${icon('help')} Como funciona</button></div></header>`;
}
function attachHeader() {
  document.querySelector('#help')!.addEventListener('click', help);
  document.querySelector('#brand')!.addEventListener('click', goHome);
}
function home() {
  app.innerHTML = `${header()}<main class="home"><div class="eyebrow">MENOS COMPLICAÇÃO. MAIS PLANILHA.</div><section class="intro"><div><h1>Sua planilha.<br><em>Sem o peso do Excel.</em></h1><p>Abra, confira, ajuste e siga o dia.<br>Um cantinho leve para suas planilhas, direto no navegador.</p></div><span class="intro-note">${icon('check')} Sem conta. Sem instalação. Sem custo.</span></section><section class="upload-zone" id="dropzone"><div><div class="file-icon">${icon('file')}</div><h2>Solte sua planilha aqui</h2><p>ou escolha um arquivo do seu computador</p><button class="primary" id="open">${icon('upload')} Abrir planilha</button></div><span class="upload-caption">Arquivos .xlsx · Processamento 100% local</span></section><div class="home-secondary"><button class="text-button" id="new">+ Criar em branco</button><span>ou</span><button class="text-button" id="demo">Explorar uma planilha de exemplo ${icon('arrow')}</button></div><section class="features"><div class="feature"><span class="feature-icon">${icon('feather')}</span><div><h3>Só o que você precisa</h3><p>Edições rápidas, várias abas e busca.<br>Sem menus para se perder.</p></div></div><div class="feature"><span class="feature-icon">${icon('lock')}</span><div><h3>Privado de verdade</h3><p>Nenhuma planilha enviada a servidores.<br>O arquivo fica no seu navegador.</p></div></div><div class="feature"><span class="feature-icon">${icon('download')}</span><div><h3>Abriu. Editou. Baixou.</h3><p>Salve uma nova cópia em .xlsx.<br>Pronta para abrir no Excel ou Numbers.</p></div></div></section><footer class="home-footer"><span>Feito para facilitar as pequenas tarefas.</span><span>LEVE POR NATUREZA ↗</span></footer></main>`;
  attachHeader();
  document.querySelector('#open')!.addEventListener('click', pickFile);
  document.querySelector('#new')!.addEventListener('click', () => newBook(false));
  document.querySelector('#demo')!.addEventListener('click', () => newBook(true));
  const dropzone = document.querySelector('#dropzone')!;
  dropzone.addEventListener('dragover', event => { event.preventDefault(); dropzone.classList.add('dragging'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragging'));
  dropzone.addEventListener('drop', event => {
    event.preventDefault(); dropzone.classList.remove('dragging');
    const file = (event as DragEvent).dataTransfer?.files[0]; if (file) void openFile(file);
  });
}
function confirmDiscard() {
  commitEditor();
  return !dirty || window.confirm('Você tem alterações que ainda não foram baixadas. Descartar e continuar?');
}
function goHome() { if (confirmDiscard()) { dirty = false; home(); } }
function pickFile() {
  const input = document.createElement('input'); input.type = 'file'; input.accept = '.xlsx';
  input.addEventListener('change', () => { if (input.files?.[0]) void openFile(input.files[0]); }); input.click();
}
async function openFile(file: File) {
  if (!/\.xlsx$/i.test(file.name)) { toast('Escolha um arquivo .xlsx. Arquivos .xls e .xlsm não são suportados.'); return; }
  if (!confirmDiscard()) return;
  await busy('Abrindo sua planilha…', async () => {
    const library = await engine();
    const nextBook = new library.Workbook();
    try { await nextBook.xlsx.load(await file.arrayBuffer()); }
    catch { throw new Error('Não foi possível abrir este XLSX. Confira se ele é válido e não está protegido por senha.'); }
    // Excel's shared-formula groups depend on a master cell. Materialize their
    // translated formulas first so editing that master cannot break its peers.
    const shared: { cell: Cell; formula: string; result: unknown }[] = [];
    nextBook.worksheets.forEach(worksheet => worksheet.eachRow(currentRow => currentRow.eachCell(cell => {
      const value = cell.value;
      if (value && typeof value === 'object' && ('sharedFormula' in value || ('shareType' in value && value.shareType === 'shared'))) {
        shared.push({ cell, formula: cell.formula, result: cell.result });
      }
    })));
    shared.forEach(({ cell, formula, result }) => { cell.value = { formula, result } as CellValue; });
    if (!nextBook.worksheets.length) nextBook.addWorksheet('Planilha 1');
    book = nextBook; filename = file.name; reset();
    sheet = book.worksheets.find(s => s.state === 'visible') ?? book.worksheets[0];
    showEditor();
  });
}
async function newBook(demo: boolean) {
  if (!confirmDiscard()) return;
  await busy('Preparando sua planilha…', async () => {
    const library = await engine(); book = new library.Workbook();
    filename = demo ? 'Planejamento mensal.xlsx' : 'Sem título.xlsx';
    sheet = book.addWorksheet(demo ? 'Despesas' : 'Planilha 1');
    if (demo) {
      const data = [
        ['PLANEJAMENTO MENSAL', '', '', '', '', ''],
        ['Um pouco de organização, mais tranquilidade.', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['Descrição', 'Categoria', 'Data', 'Valor (R$)', 'Status', 'Observações'],
        ['Aluguel', 'Moradia', '05/10/2026', 1800, 'Pago', 'Vencimento no dia 5'],
        ['Supermercado', 'Alimentação', '07/10/2026', 342.9, 'Pago', 'Compras da semana'],
        ['Internet', 'Serviços', '10/10/2026', 119.9, 'Pendente', 'Débito automático'],
        ['Café com amigos', 'Lazer', '12/10/2026', 68, 'Planejado', 'Um respiro na rotina'],
        ['Energia elétrica', 'Moradia', '15/10/2026', 156.4, 'Pendente', ''],
        ['Livraria', 'Pessoal', '18/10/2026', 89.9, 'Planejado', 'Livro do mês'],
        ['Transporte', 'Mobilidade', '20/10/2026', 120, 'Planejado', ''],
        ['', '', '', '', '', ''],
        ['Total previsto', '', '', { formula: 'SUM(D5:D11)', result: 2697.1 }, '', ''],
      ];
      data.forEach(values => sheet.addRow(values));
      sheet.getRow(1).font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF184B3E' } };
      sheet.getRow(2).font = { name: 'Calibri', size: 11, italic: true, color: { argb: 'FF86917B' } };
      sheet.getRow(4).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FF184B3E' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5EDBD' } };
      });
      sheet.getRow(13).font = { bold: true, color: { argb: 'FF184B3E' } };
      for (let r = 5; r <= 13; r++) sheet.getCell(r, 4).numFmt = '"R$" #,##0.00';
      sheet.columns.forEach(column => { column.width = 23; });
      const notes = book.addWorksheet('Notas');
      notes.addRow(['Seu espaço para anotações']);
      notes.addRow(['Dica: clique duas vezes em uma célula para editar.']);
      notes.addRow(['Cole uma tabela do Excel ou do Google Sheets com ⌘V.']);
      notes.getColumn(1).width = 65;
    }
    reset(); showEditor();
  });
}
function reset() {
  undoStack = []; redoStack = []; revision = 0; nextRevision = 0; savedRevision = 0; dirty = false;
  pendingFormula = null;
  row = col = anchorRow = anchorCol = 1; editorInput = null;
}
function showEditor() {
  app.innerHTML = `${header()}<main class="editor"><section class="doc-bar"><div class="doc-title"><button class="back" id="back" aria-label="Voltar ao início">${icon('back')}</button><div><h1 id="filename">${escape(filename)}</h1><p id="save-state">Arquivo local · Pronto para editar</p></div></div><div class="doc-actions"><button class="text-button" id="open-another">${icon('upload')} Abrir outro</button><button class="primary" id="save">${icon('download')} Baixar .xlsx</button></div></section><div class="toolbar" role="toolbar" aria-label="Formatação"><button class="tool" id="undo" title="Desfazer (⌘/Ctrl Z)" aria-label="Desfazer">${icon('undo')}</button><button class="tool" id="redo" title="Refazer (⌘/Ctrl Shift Z)" aria-label="Refazer">${icon('redo')}</button><span class="divider"></span><button class="tool" id="bold" title="Negrito" aria-label="Negrito"><b>B</b></button><button class="tool" id="italic" title="Itálico" aria-label="Itálico"><i>I</i></button><span class="divider"></span><span class="toolbar-label">FORMATO</span><button class="tool" id="format-number" title="Número com duas casas decimais">123</button><button class="tool" id="format-currency" title="Moeda (R$)">R$</button><button class="tool" id="format-percent" title="Porcentagem">%</button><button class="tool" id="format-general" title="Formato geral">Aa</button><span class="divider"></span><button class="tool" id="clear" title="Limpar conteúdo" aria-label="Limpar conteúdo">⌫</button><div class="search-wrap">${icon('search')}<input id="search" type="search" placeholder="Buscar na aba…" aria-label="Buscar na aba" /><button id="search-next" title="Próximo resultado" aria-label="Próximo resultado">↵</button></div></div><div class="formula-bar"><input class="address" id="address" value="A1" aria-label="Ir para célula" title="Digite um endereço, como D20" /><span class="fx">ƒx</span><input class="formula-input" id="formula" aria-label="Valor ou fórmula da célula" placeholder="Valor ou fórmula da célula selecionada" /></div><div class="grid" id="grid" tabindex="0" role="grid" aria-label="Planilha"><div class="grid-space" id="grid-space"></div></div><nav class="sheet-bar" id="sheets" aria-label="Abas da planilha"></nav><footer class="status-bar"><span id="selection-status"><span class="dot"></span>Pronto</span><div class="status-right"><span id="selection-stats"></span><span title="Fórmulas são preservadas; seus resultados são recalculados ao abrir no Excel.">Fórmulas: resultado do arquivo</span><span>100% local ${icon('lock')}</span></div></footer></main>`;
  grid = document.querySelector('#grid')!; space = document.querySelector('#grid-space')!;
  attachHeader();
  document.querySelector('#back')!.addEventListener('click', goHome);
  document.querySelector('#open-another')!.addEventListener('click', pickFile);
  document.querySelector('#save')!.addEventListener('click', save);
  document.querySelector('#undo')!.addEventListener('click', () => history(false));
  document.querySelector('#redo')!.addEventListener('click', () => history(true));
  document.querySelector('#bold')!.addEventListener('click', () => format('bold'));
  document.querySelector('#italic')!.addEventListener('click', () => format('italic'));
  for (const name of ['number', 'currency', 'percent', 'general']) {
    document.querySelector(`#format-${name}`)!.addEventListener('click', () => format(name));
  }
  document.querySelector('#clear')!.addEventListener('click', clearSelection);
  document.querySelector<HTMLInputElement>('#address')!.addEventListener('keydown', event => {
    if (event.key !== 'Enter') return;
    const match = (event.currentTarget as HTMLInputElement).value.trim().match(/^([A-Z]{1,3})([1-9]\d*)$/i);
    if (!match) { toast('Use um endereço como A1 ou D20.'); return; }
    const c = [...match[1].toUpperCase()].reduce((n, character) => n * 26 + character.charCodeAt(0) - 64, 0);
    const r = Number(match[2]);
    if (r > 1048576 || c > 16384) { toast('Endereço fora dos limites de uma planilha XLSX.'); return; }
    select(r, c); grid.focus();
  });
  const formula = document.querySelector<HTMLInputElement>('#formula')!;
  formula.addEventListener('input', () => { pendingFormula = { r: row, c: col, text: formula.value }; });
  formula.addEventListener('keydown', event => {
    if (event.key === 'Enter') { event.preventDefault(); commitFormula(); select(row + 1, col); grid.focus(); }
    if (event.key === 'Escape') { pendingFormula = null; updateSelection(); grid.focus(); }
  });
  formula.addEventListener('change', commitFormula);
  document.querySelector('#search')!.addEventListener('keydown', event => {
    if ((event as KeyboardEvent).key === 'Enter') search();
    if ((event as KeyboardEvent).key === 'Escape') grid.focus();
  });
  document.querySelector('#search-next')!.addEventListener('click', search);
  grid.addEventListener('scroll', () => { commitEditor(); renderGrid(); });
  grid.addEventListener('pointerdown', event => {
    const element = (event.target as HTMLElement).closest<HTMLElement>('.cell[data-row]');
    if (!element || editorInput?.contains(event.target as Node)) return;
    event.preventDefault(); commitEditor();
    const r = Number(element.dataset.row), c = Number(element.dataset.col), time = performance.now();
    const doubleClick = !event.shiftKey && lastPointer?.r === r && lastPointer.c === c && time - lastPointer.time < 450;
    lastPointer = { r, c, time };
    select(r, c, event.shiftKey, false);
    if (doubleClick) { draggingSelection = false; lastPointer = null; startEdit(); return; }
    draggingSelection = true; grid.focus();
  });
  grid.addEventListener('pointermove', event => {
    if (!draggingSelection || !(event.buttons & 1)) return;
    const element = (event.target as HTMLElement).closest<HTMLElement>('.cell[data-row]');
    if (element && (row !== Number(element.dataset.row) || col !== Number(element.dataset.col))) {
      select(Number(element.dataset.row), Number(element.dataset.col), true, false);
    }
  });
  grid.addEventListener('dblclick', event => {
    const rect = grid.getBoundingClientRect();
    if (event.clientX - rect.left > GUTTER && event.clientY - rect.top > HEADER) startEdit();
  });
  grid.addEventListener('keydown', gridKeys);
  grid.addEventListener('paste', paste);
  grid.addEventListener('copy', copy);
  refreshDimensions(); renderTabs(); renderGrid(); updateSelection();
}
function colName(c: number) {
  let name = ''; while (c) { c--; name = String.fromCharCode(65 + c % 26) + name; c = Math.floor(c / 26); } return name;
}
function bounds() { return { top: Math.min(row, anchorRow), bottom: Math.max(row, anchorRow), left: Math.min(col, anchorCol), right: Math.max(col, anchorCol) }; }
function refreshDimensions() {
  rowCount = Math.min(1048576, Math.max(100, sheet.rowCount + 30, row + 20));
  colCount = Math.min(16384, Math.max(26, sheet.columnCount + 5, col + 5));
  space.style.width = `${GUTTER + COL * colCount}px`; space.style.height = `${HEADER + ROW * rowCount}px`;
}
function valueOf(cell: Cell): unknown {
  const value = cell.value;
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;
  if (typeof value === 'object') {
    if ('richText' in value) return value.richText.map(part => part.text).join('');
    if ('formula' in value || 'sharedFormula' in value) return value.result ?? '';
    if ('hyperlink' in value) return value.text;
    if ('error' in value) return value.error;
  }
  return value;
}
function display(cell: Cell): string {
  const value = valueOf(cell);
  if (value instanceof Date) return value.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
  if (typeof value === 'boolean') return value ? 'VERDADEIRO' : 'FALSO';
  if (typeof value === 'number') {
    const fmt = cell.numFmt ?? '';
    if (fmt.includes('%')) return new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 2 }).format(value);
    const currency = fmt.includes('R$') ? 'BRL' : fmt.includes('$') ? 'USD' : fmt.includes('€') ? 'EUR' : null;
    if (currency) return new Intl.NumberFormat('pt-BR', { style: 'currency', currency }).format(value);
    if (/0\.00/.test(fmt)) return new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
    return String(value);
  }
  return String(value ?? '');
}
function raw(cell: Cell): string {
  if (cell.formula) return `=${cell.formula}`;
  const value = valueOf(cell);
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value ?? '');
}
function renderGrid() {
  if (!space || !sheet) return;
  const top = grid.scrollTop, left = grid.scrollLeft;
  const rStart = Math.max(1, Math.floor((top - HEADER) / ROW) + 1);
  const rEnd = Math.min(rowCount, Math.ceil((top + grid.clientHeight) / ROW));
  const cStart = Math.max(1, Math.floor((left - GUTTER) / COL) + 1);
  const cEnd = Math.min(colCount, Math.ceil((left + grid.clientWidth) / COL));
  const selection = bounds(); const fragment = document.createDocumentFragment();
  const make = (text: string, x: number, y: number, w: number, classes: string) => {
    const element = document.createElement('div'); element.className = `cell ${classes}`;
    element.style.cssText = `left:${x}px;top:${y}px;width:${w}px;`;
    element.textContent = text; fragment.append(element); return element;
  };
  for (let r = rStart; r <= rEnd; r++) {
    for (let c = cStart; c <= cEnd; c++) {
      // findCell avoids materializing thousands of empty workbook cells while scrolling.
      const cell = sheet.findRow(r)?.findCell(c);
      const selected = r >= selection.top && r <= selection.bottom && c >= selection.left && c <= selection.right;
      const element = make(cell ? display(cell) : '', GUTTER + (c - 1) * COL, HEADER + (r - 1) * ROW, COL,
        `${selected ? 'selected' : ''} ${r === row && c === col ? 'current' : ''} ${cell && typeof valueOf(cell) === 'number' ? 'numeric' : ''}`);
      element.dataset.row = String(r); element.dataset.col = String(c);
      element.role = 'gridcell'; element.setAttribute('aria-label', `${colName(c)}${r}${cell ? ': ' + display(cell) : ''}`);
      element.setAttribute('aria-selected', String(selected));
      if (cell) {
        element.title = raw(cell);
        if (cell.font?.bold) element.style.fontWeight = '700';
        if (cell.font?.italic) element.style.fontStyle = 'italic';
        const color = cell.font?.color?.argb;
        if (color?.length === 8) element.style.color = `#${color.slice(2)}`;
        if (cell.fill?.type === 'pattern' && cell.fill.pattern === 'solid' && !selected) {
          const bg = cell.fill.fgColor?.argb; if (bg?.length === 8) element.style.background = `#${bg.slice(2)}`;
        }
      }
    }
    make(String(r), left, HEADER + (r - 1) * ROW, GUTTER,
      `header ${r >= selection.top && r <= selection.bottom ? 'selected-header' : ''}`);
  }
  for (let c = cStart; c <= cEnd; c++) {
    make(colName(c), GUTTER + (c - 1) * COL, top, COL,
      `header ${c >= selection.left && c <= selection.right ? 'selected-header' : ''}`);
  }
  make('↘', left, top, GUTTER, 'header corner');
  space.replaceChildren(fragment);
}
function select(r: number, c: number, extend = false, scroll = true) {
  commitEditor(); row = Math.max(1, Math.min(1048576, r)); col = Math.max(1, Math.min(16384, c));
  if (!extend) { anchorRow = row; anchorCol = col; }
  refreshDimensions();
  if (scroll) {
    const x = GUTTER + (col - 1) * COL, y = HEADER + (row - 1) * ROW;
    if (x < grid.scrollLeft + GUTTER) grid.scrollLeft = x - GUTTER;
    else if (x + COL > grid.scrollLeft + grid.clientWidth) grid.scrollLeft = x + COL - grid.clientWidth;
    if (y < grid.scrollTop + HEADER) grid.scrollTop = y - HEADER;
    else if (y + ROW > grid.scrollTop + grid.clientHeight) grid.scrollTop = y + ROW - grid.clientHeight;
  }
  renderGrid(); updateSelection();
}
function updateSelection() {
  if (!document.querySelector('#formula')) return;
  const selection = bounds(); const count = (selection.bottom - selection.top + 1) * (selection.right - selection.left + 1);
  const cell = sheet.getCell(row, col);
  document.querySelector<HTMLInputElement>('#address')!.value = `${colName(col)}${row}`;
  document.querySelector<HTMLInputElement>('#formula')!.value = raw(cell);
  document.querySelector('#selection-status')!.innerHTML = `<span class="dot"></span>${count === 1 ? `${colName(col)}${row} selecionada` : `${count.toLocaleString('pt-BR')} células selecionadas`}`;
  let sum = 0, numbers = 0;
  if (count <= 10000) for (let r = selection.top; r <= selection.bottom; r++) for (let c = selection.left; c <= selection.right; c++) {
    const candidate = sheet.findRow(r)?.findCell(c); const value = candidate && valueOf(candidate);
    if (typeof value === 'number') { sum += value; numbers++; }
  }
  document.querySelector('#selection-stats')!.textContent = numbers > 1 ? `Soma: ${sum.toLocaleString('pt-BR', { maximumFractionDigits: 2 })} · Média: ${(sum / numbers).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}` : '';
  document.querySelector('#bold')!.classList.toggle('active', !!cell.font?.bold);
  document.querySelector('#italic')!.classList.toggle('active', !!cell.font?.italic);
  document.querySelector<HTMLButtonElement>('#undo')!.disabled = !undoStack.length;
  document.querySelector<HTMLButtonElement>('#redo')!.disabled = !redoStack.length;
  document.querySelector('#save-state')!.textContent = dirty ? 'Alterações locais · Baixe para salvar uma cópia' : 'Arquivo local · Pronto para editar';
}
function snapshot(cell: Cell): Snapshot {
  return { row: Number(cell.row), col: Number(cell.col), value: structuredClone(cell.value), style: structuredClone(cell.style) };
}
function mutate(cells: Cell[], action: (cell: Cell) => void) {
  const unique = [...new Map(cells.map(cell => [cell.master.address, cell.master])).values()];
  const before = unique.map(snapshot);
  unique.forEach(action);
  const after = unique.map(snapshot);
  if (JSON.stringify(before) === JSON.stringify(after)) return;
  const beforeRevision = revision; revision = ++nextRevision;
  undoStack.push({ sheetId: sheet.id, before, after, beforeRevision, afterRevision: revision });
  if (undoStack.length > 100) undoStack.shift();
  redoStack = []; dirty = true; refreshDimensions(); renderGrid(); updateSelection();
}
function selectedCells() {
  const b = bounds(); const count = (b.bottom - b.top + 1) * (b.right - b.left + 1);
  if (count > 50000) { toast('Selecione até 50 mil células por operação para manter o editor leve.'); return []; }
  const cells: Cell[] = [];
  for (let r = b.top; r <= b.bottom; r++) for (let c = b.left; c <= b.right; c++) cells.push(sheet.getCell(r, c));
  return cells;
}
function parseValue(text: string, cell: Cell): CellValue {
  if (text === '') return null;
  if (text.startsWith("'")) return text.slice(1);
  if (text.startsWith('=') && text.length > 1) return { formula: text.slice(1) };
  if (cell.value instanceof Date && /^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const date = new Date(`${text}T00:00:00Z`); if (!Number.isNaN(date.getTime())) return date;
  }
  if (/^(true|verdadeiro)$/i.test(text)) return true;
  if (/^(false|falso)$/i.test(text)) return false;
  // Preserve identifiers with leading zeroes and integers outside JS's safe range.
  if (/^[+-]?(?:0|[1-9]\d*)(?:[.,]\d+)?(?:[eE][+-]?\d+)?$/.test(text)) {
    const numeric = Number(text.replace(',', '.'));
    if (Number.isFinite(numeric) && Math.abs(numeric) <= Number.MAX_SAFE_INTEGER) return numeric;
  }
  return text;
}
function setValue(r: number, c: number, text: string) {
  const cell = sheet.getCell(r, c).master;
  if (raw(cell) === text) return;
  mutate([cell], target => { target.value = parseValue(text, target); });
}
function startEdit(initial?: string) {
  if (editorInput) return;
  anchorRow = row; anchorCol = col; renderGrid();
  const element = space.querySelector<HTMLElement>(`.cell[data-row="${row}"][data-col="${col}"]`);
  if (!element) return;
  const input = document.createElement('input'); input.className = 'cell-input';
  input.setAttribute('aria-label', `Editar ${colName(col)}${row}`);
  input.value = initial ?? raw(sheet.getCell(row, col));
  element.textContent = ''; element.classList.add('editor-cell'); element.append(input); editorInput = input;
  input.addEventListener('keydown', event => {
    event.stopPropagation();
    if (event.key === 'Escape') { editorInput = null; renderGrid(); grid.focus(); }
    else if (event.key === 'Enter' || event.key === 'Tab') {
      event.preventDefault(); const isTab = event.key === 'Tab'; const delta = event.shiftKey ? -1 : 1;
      commitEditor(); select(row + (isTab ? 0 : delta), col + (isTab ? delta : 0)); grid.focus();
    }
  });
  input.addEventListener('blur', commitEditor);
  input.focus(); if (initial === undefined) input.select();
}
function commitEditor() {
  commitFormula();
  if (!editorInput) return;
  const text = editorInput.value; editorInput = null;
  setValue(row, col, text); renderGrid();
}
function commitFormula() {
  if (!pendingFormula) return;
  const { r, c, text } = pendingFormula; pendingFormula = null;
  setValue(r, c, text);
}
function gridKeys(event: KeyboardEvent) {
  if (event.target !== grid) return;
  const command = event.ctrlKey || event.metaKey;
  if (command && event.key.toLowerCase() === 'z') { event.preventDefault(); history(event.shiftKey); return; }
  if (command && event.key.toLowerCase() === 'y') { event.preventDefault(); history(true); return; }
  if (command && event.key.toLowerCase() === 'b') { event.preventDefault(); format('bold'); return; }
  if (command && event.key.toLowerCase() === 'a') {
    event.preventDefault(); anchorRow = anchorCol = 1; select(Math.max(1, sheet.rowCount), Math.max(1, sheet.columnCount), true); return;
  }
  if (command || event.altKey) return;
  const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, event.shiftKey ? -1 : 1], Enter: [event.shiftKey ? -1 : 1, 0] };
  if (moves[event.key]) { event.preventDefault(); const [r, c] = moves[event.key]; select(row + r, col + c, event.shiftKey && event.key.startsWith('Arrow')); }
  else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearSelection(); }
  else if (event.key === 'F2') { event.preventDefault(); startEdit(); }
  else if (event.key.length === 1) { event.preventDefault(); startEdit(event.key); }
}
function clearSelection() { commitEditor(); mutate(selectedCells(), cell => { cell.value = null; }); grid.focus(); }
function format(type: string) {
  commitEditor(); const active = sheet.getCell(row, col).font;
  const turnOn = !(type === 'bold' ? active?.bold : active?.italic);
  mutate(selectedCells(), cell => {
    if (type === 'bold' || type === 'italic') cell.font = { ...cell.font, [type]: turnOn };
    else cell.numFmt = ({ number: '#,##0.00', currency: '"R$" #,##0.00', percent: '0.00%', general: 'General' })[type]!;
  }); grid.focus();
}
function history(redo: boolean) {
  commitEditor(); const change = (redo ? redoStack : undoStack).pop(); if (!change) return;
  sheet = book.getWorksheet(change.sheetId)!;
  (redo ? change.after : change.before).forEach(state => {
    const cell = sheet.getCell(state.row, state.col); cell.value = structuredClone(state.value); cell.style = structuredClone(state.style);
  });
  revision = redo ? change.afterRevision : change.beforeRevision;
  dirty = revision !== savedRevision;
  (redo ? undoStack : redoStack).push(change);
  const first = change.before[0]; if (first) { row = anchorRow = first.row; col = anchorCol = first.col; }
  renderTabs(); select(row, col); grid.focus();
}
// TSV with quoted fields: multiline cells and tabs round-trip through clipboard.
function parseTSV(text: string) {
  const rows: string[][] = []; let fields: string[] = [], field = '', quoted = false;
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' && (quoted || field === '')) {
      if (quoted && text[i + 1] === '"') { field += '"'; i++; } else quoted = !quoted;
    } else if (!quoted && (ch === '\t' || ch === '\n')) {
      fields.push(field); field = '';
      if (ch === '\n') { rows.push(fields); fields = []; }
    } else field += ch;
  }
  fields.push(field); if (fields.length > 1 || fields[0] !== '' || !rows.length) rows.push(fields);
  return rows;
}
function paste(event: ClipboardEvent) {
  if (event.target !== grid) return;
  const text = event.clipboardData?.getData('text/plain'); if (text === undefined) return;
  event.preventDefault();
  const rows = parseTSV(text); const count = rows.reduce((n, values) => n + values.length, 0);
  if (count > 50000) { toast('Cole até 50 mil células por vez.'); return; }
  if (row + rows.length - 1 > 1048576 || col + Math.max(...rows.map(values => values.length)) - 1 > 16384) { toast('A tabela ultrapassa os limites de uma planilha XLSX.'); return; }
  const edits = rows.flatMap((values, ri) => values.map((value, ci) => ({ cell: sheet.getCell(row + ri, col + ci).master, value })));
  const byAddress = new Map(edits.map(edit => [edit.cell.address, edit.value]));
  mutate(edits.map(edit => edit.cell), cell => { cell.value = parseValue(byAddress.get(cell.address)!, cell); });
  toast(`${count.toLocaleString('pt-BR')} células coladas.`);
}
function copy(event: ClipboardEvent) {
  if (event.target !== grid) return;
  const b = bounds(); if ((b.bottom - b.top + 1) * (b.right - b.left + 1) > 50000) { toast('Copie até 50 mil células por vez.'); event.preventDefault(); return; }
  const rows: string[] = [];
  for (let r = b.top; r <= b.bottom; r++) {
    const values: string[] = [];
    for (let c = b.left; c <= b.right; c++) {
      const cell = sheet.findRow(r)?.findCell(c); const text = cell ? raw(cell) : '';
      values.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    rows.push(values.join('\t'));
  }
  event.preventDefault(); event.clipboardData?.setData('text/plain', rows.join('\n'));
}
function renderTabs() {
  const nav = document.querySelector('#sheets')!; nav.replaceChildren();
  for (const worksheet of book.worksheets.filter(s => s.state === 'visible' || s.id === sheet.id)) {
    const button = document.createElement('button'); button.className = `sheet-tab ${worksheet.id === sheet.id ? 'active' : ''}`;
    button.textContent = worksheet.name; button.setAttribute('aria-current', worksheet.id === sheet.id ? 'true' : 'false');
    button.title = 'Clique para abrir · Duplo clique para renomear';
    button.addEventListener('click', () => {
      commitEditor(); sheet = worksheet; row = col = anchorRow = anchorCol = 1;
      grid.scrollTop = grid.scrollLeft = 0; refreshDimensions(); renderTabs(); renderGrid(); updateSelection();
    });
    button.addEventListener('dblclick', () => renameSheet(worksheet));
    nav.append(button);
  }
  const add = document.createElement('button'); add.className = 'add-sheet'; add.textContent = '+'; add.title = 'Adicionar aba'; add.setAttribute('aria-label', 'Adicionar aba');
  add.addEventListener('click', () => {
    commitEditor(); let index = book.worksheets.length + 1; while (book.getWorksheet(`Planilha ${index}`)) index++;
    sheet = book.addWorksheet(`Planilha ${index}`); structuralChange(); row = col = anchorRow = anchorCol = 1;
    grid.scrollTop = grid.scrollLeft = 0; refreshDimensions(); renderTabs(); renderGrid(); updateSelection();
  }); nav.append(add);
}
function structuralChange() {
  revision = ++nextRevision; dirty = true; undoStack = []; redoStack = [];
  toast('Aba atualizada. O histórico de células foi reiniciado.');
}
function renameSheet(worksheet: Worksheet) {
  const name = window.prompt('Nome da aba (até 31 caracteres):', worksheet.name)?.trim();
  if (!name || name === worksheet.name) return;
  // Renaming formula references requires an Excel formula parser; avoid broken links.
  let formulas = false;
  book.worksheets.forEach(s => s.eachRow(r => r.eachCell(c => { if (c.formula) formulas = true; })));
  if (formulas) { toast('Para preservar referências de fórmulas, renomeie as abas deste arquivo no Excel.'); return; }
  if (name.length > 31 || /[\\/*?:\[\]]/.test(name) || /^'|'$/.test(name) || book.worksheets.some(s => s.name.toLowerCase() === name.toLowerCase())) {
    toast('Nome inválido ou já utilizado. Evite \ / * ? : [ ] e use até 31 caracteres.'); return;
  }
  worksheet.name = name; structuralChange(); renderTabs(); updateSelection();
}
function search() {
  commitEditor(); const term = document.querySelector<HTMLInputElement>('#search')!.value.trim().toLocaleLowerCase('pt-BR'); if (!term) return;
  const hits: { r: number; c: number }[] = [];
  sheet.eachRow((currentRow, r) => currentRow.eachCell((cell, c) => {
    if (display(cell).toLocaleLowerCase('pt-BR').includes(term) || raw(cell).toLocaleLowerCase('pt-BR').includes(term)) hits.push({ r, c });
  }));
  if (!hits.length) { toast('Nenhum resultado nesta aba.'); return; }
  const next = hits.find(hit => hit.r > row || (hit.r === row && hit.c > col)) ?? hits[0];
  select(next.r, next.c); toast(`Resultado ${hits.indexOf(next) + 1} de ${hits.length} · ${colName(next.c)}${next.r}`);
}
async function save() {
  commitEditor();
  await busy('Preparando o arquivo para download…', async () => {
    book.calcProperties.fullCalcOnLoad = true;
    const buffer = await book.xlsx.writeBuffer();
    const blob = new Blob([new Uint8Array(buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const link = document.createElement('a');
    link.href = url; link.download = filename.replace(/\.xlsx$/i, '') + '-editado.xlsx'; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    savedRevision = revision; dirty = false; updateSelection(); toast('Download iniciado. Sua planilha original continua intacta.');
  });
}
function help() {
  const dialog = document.createElement('dialog');
  dialog.innerHTML = `<h2>Uma planilha, sem complicar.</h2><p>O Folha é um editor para conferir e fazer ajustes rápidos em arquivos <strong>.xlsx</strong>. Tudo é processado no seu navegador.</p><ul><li><strong>Editar:</strong> duplo clique, <kbd>F2</kbd> ou comece a digitar. Confirme com <kbd>Enter</kbd> ou <kbd>Tab</kbd>.</li><li><strong>Selecionar:</strong> arraste ou use <kbd>Shift</kbd> + setas. Copie e cole tabelas com <kbd>⌘/Ctrl C</kbd> e <kbd>⌘/Ctrl V</kbd>.</li><li><strong>Desfazer:</strong> <kbd>⌘/Ctrl Z</kbd>. Refazer: <kbd>⌘/Ctrl Shift Z</kbd>.</li><li><strong>Salvar:</strong> <kbd>⌘/Ctrl S</kbd> baixa uma nova cópia. A original não é sobrescrita.</li><li><strong>Números:</strong> use ponto ou vírgula decimal, sem separador de milhar. Para texto literal, comece com apóstrofo.</li></ul><p><strong>Sobre compatibilidade:</strong> valores, abas, fórmulas e estilos básicos são preservados. O Folha não calcula fórmulas: resultados do arquivo podem ficar desatualizados depois de editar. O Excel recalcula ao abrir o download. Gráficos, macros, tabelas dinâmicas e recursos avançados não têm garantia de preservação. Use uma cópia para arquivos complexos.</p><p>Não há salvamento automático. Baixe suas alterações antes de fechar a página.</p><button class="primary" id="close-help">Entendi ${icon('check')}</button>`;
  document.body.append(dialog); dialog.showModal();
  dialog.querySelector('#close-help')!.addEventListener('click', () => dialog.close());
  dialog.addEventListener('close', () => dialog.remove());
}
document.addEventListener('pointerup', () => { draggingSelection = false; });
document.addEventListener('keydown', event => {
  if (!document.querySelector('#grid') || document.querySelector('dialog[open]')) return;
  if (event.ctrlKey || event.metaKey) {
    if (event.key.toLowerCase() === 's') { event.preventDefault(); void save(); }
    if (event.key.toLowerCase() === 'f') { event.preventDefault(); document.querySelector<HTMLInputElement>('#search')!.focus(); }
    if (event.key.toLowerCase() === 'o') { event.preventDefault(); pickFile(); }
  }
});
window.addEventListener('beforeunload', event => {
  if (dirty || pendingFormula || (editorInput && editorInput.value !== raw(sheet.getCell(row, col)))) { event.preventDefault(); event.returnValue = ''; }
});
new ResizeObserver(() => { if (document.querySelector('#grid')) renderGrid(); }).observe(app);
home();

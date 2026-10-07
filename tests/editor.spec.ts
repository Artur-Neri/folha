import { test, expect, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

async function fixture() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Dados');
  sheet.addRow(['Nome', 'Valor', 'Data', 'Fórmula', 'Identificador']);
  sheet.addRow(['Café', 12.5, new Date('2026-10-07T00:00:00Z'), { formula: 'B2*2', result: 25 }, '00123']);
  sheet.getCell('B2').numFmt = '"R$" #,##0.00';
  sheet.getCell('C2').numFmt = 'dd/mm/yyyy';
  sheet.getCell('A1').font = { bold: true, color: { argb: 'FF184B3E' } };
  sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5EDBD' } };
  sheet.mergeCells('A4:B4'); sheet.getCell('A4').value = 'Mesclada';
  sheet.getCell('A6').value = { richText: [{ text: 'Olá ', font: { bold: true } }, { text: 'mundo' }] };
  sheet.getCell('B6').value = true;
  sheet.getCell('G2').value = { formula: 'B2*2', result: 25, shareType: 'shared', ref: 'G2:G4' };
  sheet.getCell('G3').value = { sharedFormula: 'G2', result: 0 };
  sheet.getCell('G4').value = { sharedFormula: 'G2', result: 0 };
  sheet.getCell('A20000').value = 'Final da planilha';
  workbook.addWorksheet('Outra aba').addRow(['Preservada']);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function upload(page: Page, buffer: Buffer) {
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Abrir planilha', exact: true }).click();
  await (await chooser).setFiles({ name: 'teste.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer });
  await expect(page.getByRole('grid')).toBeVisible();
}
async function jump(page: Page, address: string) {
  await page.getByRole('textbox', { name: 'Ir para célula' }).fill(address);
  await page.getByRole('textbox', { name: 'Ir para célula' }).press('Enter');
}
async function edit(page: Page, address: string, value: string) {
  await jump(page, address);
  await page.getByRole('grid').press('F2');
  await page.getByRole('textbox', { name: `Editar ${address}`, exact: true }).fill(value);
  await page.getByRole('textbox', { name: `Editar ${address}`, exact: true }).press('Enter');
}
async function downloadBook(page: Page) {
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Baixar .xlsx' }).click();
  const result = await download;
  expect(result.suggestedFilename()).toMatch(/-editado\.xlsx$/);
  const workbook = new ExcelJS.Workbook();
  const buffer = await readFile((await result.path())!);
  const zip = await JSZip.loadAsync(buffer);
  expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('fullCalcOnLoad="1"');
  await workbook.xlsx.load(buffer);
  return workbook;
}

test('importa, edita e exporta preservando valores, fórmulas, estilos e abas', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/'); await upload(page, await fixture());
  await expect(page.locator('.cell[data-row="2"][data-col="1"]')).toHaveText('Café');
  await expect(page.locator('.cell[data-row="2"][data-col="2"]')).toHaveText(/12,50/);
  await expect(page.locator('.cell[data-row="2"][data-col="3"]')).toHaveText('07/10/2026');
  await page.locator('.cell[data-row="2"][data-col="1"]').dblclick();
  await expect(page.getByRole('textbox', { name: 'Editar A2', exact: true })).toBeVisible();
  await page.getByRole('textbox', { name: 'Editar A2', exact: true }).fill('Café especial');
  await page.getByRole('textbox', { name: 'Editar A2', exact: true }).press('Enter');
  await edit(page, 'B2', '19,90');
  await page.getByRole('button', { name: 'Desfazer', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Valor ou fórmula da célula' })).toHaveValue('12.5');
  await page.getByRole('button', { name: 'Refazer', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Valor ou fórmula da célula' })).toHaveValue('19.9');
  await edit(page, 'G2', '7');
  await jump(page, 'A20000');
  await expect(page.locator('.cell[data-row="20000"][data-col="1"]')).toHaveText('Final da planilha');
  expect(await page.locator('.cell').count()).toBeLessThan(600);
  await page.getByRole('button', { name: 'Outra aba', exact: true }).click();
  await expect(page.locator('.cell[data-row="1"][data-col="1"]')).toHaveText('Preservada');
  const output = await downloadBook(page); const sheet = output.getWorksheet('Dados')!;
  expect(sheet.getCell('A2').value).toBe('Café especial');
  expect(sheet.getCell('B2').value).toBe(19.9);
  expect(sheet.getCell('B2').numFmt).toBe('"R$" #,##0.00');
  expect(sheet.getCell('C2').value).toEqual(new Date('2026-10-07T00:00:00Z'));
  expect(sheet.getCell('D2').formula).toBe('B2*2');
  expect(sheet.getCell('E2').value).toBe('00123');
  expect(sheet.getCell('A1').font.bold).toBe(true);
  expect(sheet.getCell('A1').fill).toMatchObject({ fgColor: { argb: 'FFE5EDBD' } });
  expect(sheet.getCell('B4').master.address).toBe('A4');
  expect(sheet.getCell('A6').value).toMatchObject({ richText: [{ text: 'Olá ' }, { text: 'mundo' }] });
  expect(sheet.getCell('B6').value).toBe(true);
  expect(sheet.getCell('G2').value).toBe(7);
  expect(sheet.getCell('G3').formula).toBe('B3*2');
  expect(sheet.getCell('G4').formula).toBe('B4*2');
  expect(output.getWorksheet('Outra aba')!.getCell('A1').value).toBe('Preservada');
  expect(errors).toEqual([]);
});

test('cola tabelas, formata seleção, busca e salva edição da barra de fórmula', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: 'Criar em branco' }).click();
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByRole('grid').evaluate(element => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'Produto\tPreço\nLivro\t42,90\n"Linha 1\nLinha 2"\t00123\n');
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData, bubbles: true }));
  });
  await jump(page, 'A1'); await page.getByRole('grid').press('Shift+ArrowRight');
  await page.getByRole('button', { name: 'Negrito', exact: true }).click();
  await page.getByRole('searchbox', { name: 'Buscar na aba' }).fill('Livro');
  await page.getByRole('searchbox', { name: 'Buscar na aba' }).press('Enter');
  await expect(page.getByRole('textbox', { name: 'Ir para célula' })).toHaveValue('A2');
  await page.getByRole('textbox', { name: 'Valor ou fórmula da célula' }).fill('Livro editado');
  // Save while the formula input still has focus, exercising pending-edit commit.
  await page.getByRole('textbox', { name: 'Valor ou fórmula da célula' }).press('Control+s');
  await expect(page.locator('#save-state')).toHaveText('Arquivo local · Pronto para editar');
  const output = await downloadBook(page); const sheet = output.worksheets[0];
  expect(sheet.getCell('A1').font.bold).toBe(true);
  expect(sheet.getCell('B1').font.bold).toBe(true);
  expect(sheet.getCell('A2').value).toBe('Livro editado');
  expect(sheet.getCell('B2').value).toBe(42.9);
  expect(sheet.getCell('A3').value).toBe('Linha 1\nLinha 2');
  expect(sheet.getCell('B3').value).toBe('00123');
  await page.getByRole('button', { name: 'Adicionar aba' }).click();
  await expect(page.getByRole('button', { name: 'Planilha 2', exact: true })).toBeVisible();
});

test('versão portátil abre via file://, edita e baixa sem nenhuma requisição de rede', async ({ page }) => {
  const network: string[] = [], errors: string[] = [];
  page.on('request', request => { if (/^https?:/.test(request.url())) network.push(request.url()); });
  page.on('pageerror', error => errors.push(error.message));
  await page.goto(`file://${resolve('portable/index.html')}`);
  await page.getByRole('button', { name: 'Explorar uma planilha de exemplo' }).click();
  await expect(page.getByRole('grid')).toBeVisible();
  await edit(page, 'A5', 'Aluguel atualizado');
  const output = await downloadBook(page);
  expect(output.getWorksheet('Despesas')!.getCell('A5').value).toBe('Aluguel atualizado');
  expect(network).toEqual([]); expect(errors).toEqual([]);
});

test('layout cabe em tela pequena e mantém o editor utilizável', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  await page.getByRole('button', { name: 'Criar em branco' }).click();
  await expect(page.getByRole('grid')).toBeVisible();
  await edit(page, 'A1', 'Pequena tela');
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  expect((await page.getByRole('grid').boundingBox())!.height).toBeGreaterThan(300);
});

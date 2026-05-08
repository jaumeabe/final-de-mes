import { neon } from '@neondatabase/serverless';
import ExcelJS from 'exceljs';

const sql = neon(process.env.DATABASE_URL);

const ADMIN_USER = process.env.ADMIN_USER || 'admin';

const GRANJAS_MADRES = [
  'PORCELIBOR','CASTELLNOU','PI','SISALLAR 1','SENTERADA','SISALLAR 3',
  'CASERIO','MASIA','GRANADELLA','FAYON ABUELA','INDUSTRIAL','CARTUJA 2',
  'MARRUGAT','SINOGA','PASCUALET','NOVIPORCI','LES SERRES','ALFUSPI',
  'FUSTERO','GUINEU','SANTA ROSA','GIRALT','EL SOLER','VENDRELL',
  'PORDALL','RUBIO','ESCODA','CARTUJA 1','PORDECONA','SISALLAR 4',
];

const SECTIONS = [
  {
    title: 'INVENTARIO',
    fields: [
      { key: 'num_machos',              label: 'Número de machos' },
      { key: 'num_cerdas',              label: 'Número de cerdas' },
      { key: 'num_primerizas',          label: 'Número de primerizas' },
      { key: 'total_cerdas_primerizas', label: 'Total' },
      { key: 'lechones_maternidad',     label: 'Lechones maternidad' },
      { key: 'lechones_destete',        label: 'Lechones destete' },
    ],
  },
  {
    title: 'ENTRADAS',
    fields: [
      { key: 'primerizas_entradas',  label: 'Nº primerizas entradas' },
      { key: 'acumulado_primerizas', label: 'Acumulado primerizas' },
      { key: 'machos_entrados',      label: 'Nº machos entrados' },
    ],
  },
  {
    title: 'VENTAS',
    fields: [
      { key: 'venta_cerdas',             label: 'Venta cerdas' },
      { key: 'acum_venta_cerdas',        label: 'Acumulado venta cerdas' },
      { key: 'venta_primerizas',         label: 'Venta primerizas' },
      { key: 'acum_venta_primerizas',    label: 'Acumulado venta primerizas' },
      { key: 'venta_lechones_parideras', label: 'Venta lechones parideras' },
      { key: 'acum_venta_parideras',     label: 'Acumulado venta parideras' },
      { key: 'venta_lechones_destete',   label: 'Venta lechones destete' },
      { key: 'acum_venta_destete',       label: 'Acumulado venta destete' },
      { key: 'venta_tostones',           label: 'Venta tostones' },
      { key: 'acum_venta_tostones',      label: 'Acumulado venta tostones' },
    ],
  },
  {
    title: 'PARTOS Y DESTETE',
    fields: [
      { key: 'num_partos',              label: 'Número de partos' },
      { key: 'lechones_nacidos_vivos',  label: 'Lechones nacidos vivos' },
      { key: 'prom_nacidos_vivos',      label: 'Promedio nacidos vivos' },
      { key: 'cerdas_destetadas',       label: 'Cerdas destetadas' },
      { key: 'lechones_destetados',     label: 'Lechones destetados' },
      { key: 'prom_destetados',         label: 'Promedio destetados' },
      { key: 'entrados_destete_propio', label: 'Entrados en destete propio' },
    ],
  },
  {
    title: 'CUBRICIONES',
    fields: [
      { key: 'cubriciones_totales',    label: 'Cubriciones totales' },
      { key: 'cubriciones_conflictivo',label: 'Cubriciones origen conflictivo' },
      { key: 'cubiertas_primera_vez',  label: 'Cubiertas por primera vez' },
    ],
  },
  {
    title: 'MUERTES Y BAJAS',
    fields: [
      { key: 'muerte_cerdas',         label: 'Muerte cerdas' },
      { key: 'acum_muerte_cerdas',    label: 'Acumulado muerte cerdas' },
      { key: 'muerte_primerizas',     label: 'Muerte primerizas' },
      { key: 'acum_muerte_primerizas',label: 'Acumulado muerte primerizas' },
      { key: 'bajas_destete',         label: 'Bajas destete' },
      { key: 'bajas_parideras',       label: 'Bajas parideras al mes' },
    ],
  },
];

function getMonthLabel(mesKey) {
  const [y, m] = mesKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

async function generateResumenExcel(rows, mesKey) {
  // Mapear datos por granja
  const dataByGranja = {};
  for (const row of rows) {
    dataByGranja[row.granja] = row.data;
  }

  // Solo granjas que han reportado, ordenadas según GRANJAS_MADRES
  const granjas = GRANJAS_MADRES.filter(g => dataByGranja[g]);
  const numGranjas = granjas.length;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Final de Mes');

  // Estilos
  const headerFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const sectionFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const altFill     = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F9FC' } };
  const borderThin  = {
    top:    { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left:   { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right:  { style: 'thin', color: { argb: 'FFB0B0B0' } },
  };
  const whiteFont    = { bold: true, size: 9, name: 'Arial', color: { argb: 'FFFFFFFF' } };
  const labelFont    = { size: 9, name: 'Arial', color: { argb: 'FF2C4A6E' } };
  const numFont      = { size: 9, name: 'Arial' };
  const titleFont    = { bold: true, size: 13, name: 'Arial', color: { argb: 'FF1A3A5C' } };
  const subtitleFont = { size: 10, name: 'Arial', color: { argb: 'FF4A6080' } };

  // Anchos de columna
  ws.getColumn(1).width = 28;
  for (let i = 2; i <= numGranjas + 1; i++) {
    ws.getColumn(i).width = 12;
  }

  const lastCol = numGranjas + 1;

  // ── Fila 1: Título ──────────────────────────────────────
  const titleRow = ws.getRow(1);
  titleRow.height = 24;
  const titleCell = titleRow.getCell(1);
  titleCell.value = 'Final de Mes — resumen';
  titleCell.font = titleFont;
  titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
  if (numGranjas > 0) ws.mergeCells(1, 1, 1, lastCol);

  // ── Fila 2: Mes ─────────────────────────────────────────
  const mesRow = ws.getRow(2);
  mesRow.height = 16;
  const mesCell = mesRow.getCell(1);
  mesCell.value = `Mes: ${getMonthLabel(mesKey)}`;
  mesCell.font = subtitleFont;
  mesCell.alignment = { horizontal: 'left', vertical: 'middle' };
  if (numGranjas > 0) ws.mergeCells(2, 1, 2, lastCol);

  // ── Fila 3: vacía ───────────────────────────────────────
  ws.getRow(3).height = 6;

  // ── Fila 4: Cabecera de granjas ──────────────────────────
  const headerRow = ws.getRow(4);
  headerRow.height = 20;
  headerRow.getCell(1).value = 'INVENTARIO'; // placeholder, se sobrescribe
  granjas.forEach((g, i) => {
    const cell = headerRow.getCell(i + 2);
    cell.value = g;
    cell.font = whiteFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = borderThin;
  });
  // Celda vacía A4
  const a4 = headerRow.getCell(1);
  a4.fill = headerFill;
  a4.border = borderThin;
  a4.value = '';

  // ── Datos por sección ───────────────────────────────────
  let rowIdx = 5;
  let sectionIdx = 0;

  for (const section of SECTIONS) {
    // Fila de sección
    const secRow = ws.getRow(rowIdx);
    secRow.height = 16;
    const secCell = secRow.getCell(1);
    secCell.value = section.title;
    secCell.font = whiteFont;
    secCell.fill = sectionFill;
    secCell.alignment = { vertical: 'middle' };
    secCell.border = borderThin;
    for (let c = 2; c <= lastCol; c++) {
      const cell = secRow.getCell(c);
      cell.fill = sectionFill;
      cell.border = borderThin;
    }
    if (numGranjas > 0) ws.mergeCells(rowIdx, 1, rowIdx, lastCol);
    rowIdx++;

    // Filas de datos
    for (let fi = 0; fi < section.fields.length; fi++) {
      const field = section.fields[fi];
      const dataRow = ws.getRow(rowIdx);
      dataRow.height = 14;

      // Fondo alternado
      const bg = fi % 2 === 1 ? altFill : null;

      const labelCell = dataRow.getCell(1);
      labelCell.value = field.label;
      labelCell.font = { ...labelFont, bold: false };
      labelCell.alignment = { vertical: 'middle' };
      labelCell.border = borderThin;
      if (bg) labelCell.fill = bg;

      granjas.forEach((g, gi) => {
        const cell = dataRow.getCell(gi + 2);
        const raw = dataByGranja[g]?.[field.key];
        const isDecimal = field.key.startsWith('prom_');
        if (raw !== undefined && raw !== null && raw !== '') {
          cell.value = isDecimal ? parseFloat(raw) : (parseInt(raw, 10) || 0);
          if (isDecimal) {
            cell.numFmt = '0.00';
          } else {
            cell.numFmt = '#,##0';
          }
        } else {
          cell.value = null;
        }
        cell.font = numFont;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = borderThin;
        if (bg) cell.fill = bg;
      });

      rowIdx++;
    }
    sectionIdx++;
  }

  // Inmovilizar la fila de cabecera de granjas (fila 4)
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4, topLeftCell: 'B5' }];

  ws.pageSetup = {
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    orientation: 'landscape', paperSize: 9,
  };

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export default async function handler(req, res) {
  // ── Auth ────────────────────────────────────────────────
  const { user, pass, mes } = req.query;
  const validUser = ADMIN_USER;
  const validPass = process.env.ADMIN_PASS || process.env.CRON_SECRET;

  if (!user || !pass || user !== validUser || pass !== validPass) {
    return res.status(401).json({ error: 'Credenciales incorrectas' });
  }

  if (!mes || !/^\d{4}-\d{2}$/.test(mes)) {
    return res.status(400).json({ error: 'Parámetro mes no válido (YYYY-MM)' });
  }

  // ── Consultar DB ────────────────────────────────────────
  let rows;
  try {
    rows = await sql`
      SELECT granja, data
      FROM inventario
      WHERE mes = ${mes} AND granja = ANY(${GRANJAS_MADRES})
      ORDER BY granja`;
  } catch (dbErr) {
    console.error('Error DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }

  if (!rows || rows.length === 0) {
    return res.status(200).json({ error: `No hay datos para ${mes}` });
  }

  // ── Generar Excel ───────────────────────────────────────
  try {
    const buffer = await generateResumenExcel(rows, mes);
    const filename = `Resumen_FinalDeMes_${mes}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Error generando Excel:', err);
    return res.status(500).json({ error: 'Error al generar el Excel' });
  }
}

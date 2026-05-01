import { Resend } from 'resend';
import { neon } from '@neondatabase/serverless';
import ExcelJS from 'exceljs';

const resend = new Resend(process.env.RESEND_API_KEY);
const sql = neon(process.env.DATABASE_URL);

const DESTINATARIOS = (process.env.EMAIL_TO || '').split(',').map(e => e.trim()).filter(Boolean);

const GRANJAS_MADRES = [
  'PORCELIBOR','CASTELLNOU','PI','SISALLAR 1','SENTERADA','SISALLAR 3',
  'CASERIO','MASIA','GRANADELLA','FAYON ABUELA','INDUSTRIAL','CARTUJA 2',
  'MARRUGAT','SINOGA','PASCUALET','NOVIPORCI','LES SERRES','ALFUSPI',
  'FUSTERO','GUINEU','SANTA ROSA','GIRALT','EL SOLER','VENDRELL',
  'PORDALL','ESCODA','CARTUJA 1','PORDECONA','SISALLAR 4',
];

const GRANJAS_DESTETE = [
  'CEREALS','MAIALS','LLOBET','INGLES','ZAIDIN','JAUMANDREU','BORGES 1','RUBIO',
];

// Granjas temporalmente sin animales: no cuentan para el trigger de EXISTENCIAS
const GRANJAS_PAUSADAS = ['ALFUSPI'];

const GRANJAS_MADRES_ACTIVAS = GRANJAS_MADRES.filter(g => !GRANJAS_PAUSADAS.includes(g));
const TOTAL_GRANJAS = GRANJAS_MADRES_ACTIVAS.length;

const GRANJAS_VALIDAS = [...GRANJAS_MADRES, ...GRANJAS_DESTETE];

const LABELS = {
  num_machos:               'Número de machos',
  num_cerdas:               'Número de cerdas',
  num_primerizas:           'Número de primerizas',
  total_cerdas_primerizas:  'Total',
  lechones_maternidad:      'Lechones maternidad',
  lechones_destete:         'Lechones destete',
  primerizas_entradas:      'Nº primerizas entradas',
  acumulado_primerizas:     'Acumulado primerizas',
  machos_entrados:          'Nº machos entrados',
  venta_cerdas:             'Venta cerdas',
  acum_venta_cerdas:        'Acumulado venta cerdas',
  venta_primerizas:         'Venta primerizas',
  acum_venta_primerizas:    'Acumulado venta primerizas',
  venta_lechones_parideras: 'Venta lechones parideras',
  acum_venta_parideras:     'Acumulado venta parideras',
  venta_lechones_destete:   'Venta lechones destete',
  acum_venta_destete:       'Acumulado venta destete',
  venta_tostones:           'Venta tostones',
  acum_venta_tostones:      'Acumulado venta tostones',
  muerte_cerdas:            'Muerte cerdas',
  acum_muerte_cerdas:       'Acumulado muerte cerdas',
  muerte_primerizas:        'Muerte primerizas',
  acum_muerte_primerizas:   'Acumulado muerte primerizas',
  num_partos:               'Número de partos',
  lechones_nacidos_vivos:   'Lechones nacidos vivos',
  prom_nacidos_vivos:       'Promedio nacidos vivos',
  cerdas_destetadas:        'Cerdas destetadas',
  lechones_destetados:      'Lechones destetados',
  prom_destetados:          'Promedio destetados',
  entrados_destete_propio:  'Entrados en destete propio',
  cubriciones_totales:      'Cubriciones totales',
  cubriciones_conflictivo:  'Cubriciones origen conflictivo',
  cubiertas_primera_vez:    'Cubiertas por primera vez',
  bajas_destete:            'Bajas destete',
  bajas_parideras:          'Bajas parideras al mes',
};

const SECTIONS = [
  { title: 'INVENTARIO',       fields: ['num_machos','num_cerdas','num_primerizas','total_cerdas_primerizas','lechones_maternidad','lechones_destete'] },
  { title: 'ENTRADAS',         fields: ['primerizas_entradas','acumulado_primerizas','machos_entrados'] },
  { title: 'VENTAS',           fields: ['venta_cerdas','acum_venta_cerdas','venta_primerizas','acum_venta_primerizas','venta_lechones_parideras','acum_venta_parideras','venta_lechones_destete','acum_venta_destete','venta_tostones','acum_venta_tostones'] },
  { title: 'PARTOS Y DESTETE', fields: ['num_partos','lechones_nacidos_vivos','prom_nacidos_vivos','cerdas_destetadas','lechones_destetados','prom_destetados','entrados_destete_propio'] },
  { title: 'CUBRICIONES',      fields: ['cubriciones_totales','cubriciones_conflictivo','cubiertas_primera_vez'] },
  { title: 'MUERTES Y BAJAS',  fields: ['muerte_cerdas','acum_muerte_cerdas','muerte_primerizas','acum_muerte_primerizas','bajas_destete','bajas_parideras'] },
];

const DESTETE_LABELS = {
  existencias: 'Existencias inicio mes',
  salidas:     'Salidas',
  entradas:    'Entradas',
  bajas:       'Bajas',
  final_mes:   'Final de mes',
};

const DESTETE_FIELDS = ['existencias','salidas','entradas','bajas','final_mes'];

// Campos del Excel EXISTENCIAS (filas 3-7 del Excel original)
const EXISTENCIAS_FIELDS = [
  { key: 'num_machos', label: 'MACHOS' },
  { key: 'num_cerdas', label: 'CERDAS' },
  { key: 'num_primerizas', label: 'PRIMERIZAS' },
  { key: 'lechones_maternidad', label: 'LECHONES MATERNIDAD' },
  { key: 'lechones_destete', label: 'LECHONES DESTETE' },
];

function getMonthColumn(mesKey) {
  const month = parseInt(mesKey.split('-')[1], 10);
  return month + 1;
}

function getMonthLabel(mesKey) {
  const [y, m] = mesKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

function buildHtmlMadres(data) {
  const sectionRows = SECTIONS.map(({ title, fields }) => {
    const rows = fields.map(f => `
      <tr>
        <td style="padding:6px 12px;color:#4a6080;font-size:13px;">${LABELS[f]}</td>
        <td style="padding:6px 12px;text-align:right;font-weight:600;color:#1a3a5c;font-size:13px;">${data[f] || '—'}</td>
      </tr>`).join('');
    return `
      <tr>
        <td colspan="2" style="background:#1a3a5c;color:#fff;font-size:11px;font-weight:700;letter-spacing:0.06em;padding:7px 12px;text-transform:uppercase;">${title}</td>
      </tr>
      ${rows}`;
  }).join('');

  return `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:580px;margin:0 auto;background:#fff;border-radius:10px;overflow:hidden;border:1px solid #dde3ec;">
    <div style="background:#1a3a5c;padding:24px 32px;">
      <h1 style="margin:0;color:#fff;font-size:20px;letter-spacing:0.04em;text-transform:uppercase;">Final de Mes — Granja</h1>
      <p style="margin:6px 0 0;color:#a8c4e0;font-size:14px;">${data.granja} &nbsp;|&nbsp; ${data.mes}</p>
    </div>
    <div style="padding:8px 20px 24px;">
      <table style="width:100%;border-collapse:collapse;">
        ${sectionRows}
      </table>
    </div>
    <div style="background:#f0f2f5;padding:12px 24px;font-size:11px;color:#7a8fa8;text-align:center;">
      Enviado automáticamente desde el formulario Final de Mes · Premier Pigs
    </div>
  </div>`;
}

function buildHtmlDestete(data) {
  const rows = DESTETE_FIELDS.map(f => `
    <tr>
      <td style="padding:8px 14px;color:#4a6080;font-size:13px;">${DESTETE_LABELS[f]}</td>
      <td style="padding:8px 14px;text-align:right;font-weight:600;color:#1a3a5c;font-size:13px;">${data[f] || '—'}</td>
    </tr>`).join('');

  return `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:520px;margin:0 auto;background:#fff;border-radius:10px;overflow:hidden;border:1px solid #dde3ec;">
    <div style="background:#1a3a5c;padding:24px 32px;">
      <h1 style="margin:0;color:#fff;font-size:20px;letter-spacing:0.04em;text-transform:uppercase;">Stock Destete — Granja</h1>
      <p style="margin:6px 0 0;color:#a8c4e0;font-size:14px;">${data.granja} &nbsp;|&nbsp; ${data.mes}</p>
    </div>
    <div style="padding:14px 20px 24px;">
      <table style="width:100%;border-collapse:collapse;">
        <tr>
          <td colspan="2" style="background:#1a3a5c;color:#fff;font-size:11px;font-weight:700;letter-spacing:0.06em;padding:7px 14px;text-transform:uppercase;">Inventario Destete</td>
        </tr>
        ${rows}
      </table>
    </div>
    <div style="background:#f0f2f5;padding:12px 24px;font-size:11px;color:#7a8fa8;text-align:center;">
      Enviado automáticamente desde el formulario Final de Mes · Premier Pigs
    </div>
  </div>`;
}

async function generateExcelMadres(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Final de Mes');

  const headerFont = { bold: true, size: 8, name: 'Arial', color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const labelFont = { size: 8, name: 'Arial' };
  const valueFont = { bold: true, size: 8, name: 'Arial' };
  const borderThin = {
    top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  };

  ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1, orientation: 'portrait' };
  ws.properties.defaultRowHeight = 13;

  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 14;

  ws.getRow(1).getCell(1).value = `Final de Mes — ${data.granja}`;
  ws.getRow(1).getCell(1).font = { bold: true, size: 11, name: 'Arial', color: { argb: 'FF1A3A5C' } };
  ws.mergeCells('A1:B1');
  ws.getRow(1).height = 20;

  ws.getRow(2).getCell(1).value = `Mes: ${data.mes}`;
  ws.getRow(2).getCell(1).font = { size: 9, name: 'Arial', color: { argb: 'FF4A6080' } };
  ws.getRow(2).height = 15;

  let rowIdx = 3;
  for (const section of SECTIONS) {
    const sectionRow = ws.getRow(rowIdx);
    sectionRow.getCell(1).value = section.title;
    sectionRow.getCell(1).font = headerFont;
    sectionRow.getCell(1).fill = headerFill;
    sectionRow.getCell(2).fill = headerFill;
    sectionRow.getCell(1).border = borderThin;
    sectionRow.getCell(2).border = borderThin;
    rowIdx++;

    for (const field of section.fields) {
      const r = ws.getRow(rowIdx);
      r.getCell(1).value = LABELS[field];
      r.getCell(1).font = labelFont;
      r.getCell(1).border = borderThin;
      r.getCell(2).value = data[field] ? Number(data[field]) || data[field] : '';
      r.getCell(2).font = valueFont;
      r.getCell(2).alignment = { horizontal: 'center' };
      r.getCell(2).border = borderThin;
      rowIdx++;
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

async function generateExcelDestete(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Stock Destete');

  const headerFont = { bold: true, size: 9, name: 'Arial', color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const labelFont = { size: 9, name: 'Arial' };
  const valueFont = { bold: true, size: 9, name: 'Arial' };
  const borderThin = {
    top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  };

  ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1, orientation: 'portrait' };
  ws.properties.defaultRowHeight = 15;

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 14;
  ws.getColumn(6).width = 14;

  ws.getRow(1).getCell(1).value = `Stock Destete — ${data.granja}`;
  ws.getRow(1).getCell(1).font = { bold: true, size: 12, name: 'Arial', color: { argb: 'FF1A3A5C' } };
  ws.mergeCells('A1:F1');
  ws.getRow(1).height = 22;

  ws.getRow(2).getCell(1).value = `Mes: ${data.mes}`;
  ws.getRow(2).getCell(1).font = { size: 10, name: 'Arial', color: { argb: 'FF4A6080' } };
  ws.mergeCells('A2:F2');
  ws.getRow(2).height = 16;

  // Cabecera tipo Excel histórico: MES | EXISTENCIAS | SALIDAS | ENTRADAS | BAJAS | FINAL DE MES
  const headers = ['MES','EXISTENCIAS','SALIDAS','ENTRADAS','BAJAS','FINAL DE MES'];
  const headerRow = ws.getRow(4);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = borderThin;
  });
  headerRow.height = 18;

  const dataRow = ws.getRow(5);
  dataRow.getCell(1).value = data.mes;
  dataRow.getCell(2).value = Number(data.existencias) || 0;
  dataRow.getCell(3).value = Number(data.salidas) || 0;
  dataRow.getCell(4).value = Number(data.entradas) || 0;
  dataRow.getCell(5).value = Number(data.bajas) || 0;
  dataRow.getCell(6).value = Number(data.final_mes) || 0;
  for (let c = 1; c <= 6; c++) {
    dataRow.getCell(c).font = valueFont;
    dataRow.getCell(c).alignment = { horizontal: 'center' };
    dataRow.getCell(c).border = borderThin;
    if (c >= 2) dataRow.getCell(c).numFmt = '#,##0';
  }
  dataRow.height = 18;

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

async function generateExistencias(madresRows, desteteRows, mesKey) {
  const wb = new ExcelJS.Workbook();
  const year = mesKey.split('-')[0];
  const ws = wb.addWorksheet(year);
  const col = getMonthColumn(mesKey);

  const meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  const headerFont = { bold: true, size: 9, name: 'Arial', color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const labelFont = { bold: true, size: 9, name: 'Arial' };
  const dataFont = { size: 9, name: 'Arial' };
  const borderThin = {
    top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  };

  ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1, orientation: 'landscape' };
  ws.properties.defaultRowHeight = 15;

  ws.getColumn(1).width = 20;
  for (let i = 2; i <= 13; i++) ws.getColumn(i).width = 10;

  const titleRow = ws.getRow(1);
  titleRow.getCell(2).value = `EXISTENCIAS ${year}`;
  titleRow.getCell(2).font = { bold: true, size: 11, name: 'Arial', color: { argb: 'FF1A3A5C' } };
  ws.mergeCells('B1:M1');
  titleRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
  titleRow.height = 20;

  const monthRow = ws.getRow(2);
  for (let i = 0; i < 12; i++) {
    const cell = monthRow.getCell(i + 2);
    cell.value = meses[i];
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = borderThin;
  }
  monthRow.height = 18;

  const totals = {};
  for (const field of EXISTENCIAS_FIELDS) {
    totals[field.key] = 0;
    for (const row of madresRows) {
      const val = Number(row.data[field.key]) || 0;
      totals[field.key] += val;
    }
  }

  EXISTENCIAS_FIELDS.forEach((field, idx) => {
    const row = ws.getRow(3 + idx);
    row.getCell(1).value = field.label;
    row.getCell(1).font = labelFont;
    row.getCell(1).border = borderThin;

    const cell = row.getCell(col);
    cell.value = totals[field.key];
    cell.font = dataFont;
    cell.alignment = { horizontal: 'center' };
    cell.border = borderThin;
    cell.numFmt = '#,##0';

    for (let c = 2; c <= 13; c++) {
      row.getCell(c).border = borderThin;
    }
  });

  // Fila adicional: DESTETE EXTERNO (suma de final_mes de las granjas destete)
  const desteteTotal = (desteteRows || []).reduce(
    (sum, row) => sum + (Number(row.data.final_mes) || 0), 0
  );
  const desteteRow = ws.getRow(3 + EXISTENCIAS_FIELDS.length);
  desteteRow.getCell(1).value = 'DESTETE EXTERNO';
  desteteRow.getCell(1).font = labelFont;
  desteteRow.getCell(1).border = borderThin;
  const desteteCell = desteteRow.getCell(col);
  desteteCell.value = desteteTotal;
  desteteCell.font = dataFont;
  desteteCell.alignment = { horizontal: 'center' };
  desteteCell.border = borderThin;
  desteteCell.numFmt = '#,##0';
  for (let c = 2; c <= 13; c++) {
    desteteRow.getCell(c).border = borderThin;
  }

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const data = req.body;

  if (!data.granja || !data.mes) {
    return res.status(400).json({ error: 'Faltan campos obligatorios' });
  }

  if (!GRANJAS_VALIDAS.includes(data.granja)) {
    return res.status(400).json({ error: 'Granja no válida' });
  }

  if (!/^\d{4}-\d{2}$/.test(data.mes)) {
    return res.status(400).json({ error: 'Formato de mes no válido' });
  }

  const esDestete = GRANJAS_DESTETE.includes(data.granja);
  data.tipo = esDestete ? 'destete' : 'madres';

  if (esDestete) {
    // Calculado: bajas = existencias - salidas + entradas - final_mes
    data.bajas = (Number(data.existencias) || 0)
               - (Number(data.salidas) || 0)
               + (Number(data.entradas) || 0)
               - (Number(data.final_mes) || 0);
  } else {
    data.total_cerdas_primerizas = (Number(data.num_cerdas) || 0) + (Number(data.num_primerizas) || 0);
    data.entrados_destete_propio = (Number(data.lechones_destetados) || 0) - (Number(data.venta_lechones_parideras) || 0);
  }

  // Guardar en PostgreSQL
  try {
    await sql`
      CREATE TABLE IF NOT EXISTS inventario (
        id SERIAL PRIMARY KEY,
        granja TEXT NOT NULL,
        mes TEXT NOT NULL,
        data JSONB NOT NULL,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(granja, mes)
      )`;
    await sql`
      INSERT INTO inventario (granja, mes, data)
      VALUES (${data.granja}, ${data.mes}, ${JSON.stringify(data)})
      ON CONFLICT (granja, mes)
      DO UPDATE SET data = ${JSON.stringify(data)}, created_at = NOW()`;
  } catch (dbErr) {
    console.error('Error guardando en DB:', dbErr);
    return res.status(500).json({ error: 'Error al guardar los datos' });
  }

  // Generar Excel individual
  let excelIndividual;
  try {
    excelIndividual = esDestete ? await generateExcelDestete(data) : await generateExcelMadres(data);
  } catch (xlErr) {
    console.error('Error generando Excel individual:', xlErr);
  }

  // Enviar email individual
  try {
    const subject = esDestete
      ? `Stock Destete — ${data.granja} — ${data.mes}`
      : `Final de Mes — ${data.granja} — ${data.mes}`;
    const html = esDestete ? buildHtmlDestete(data) : buildHtmlMadres(data);
    const filename = esDestete
      ? `StockDestete_${data.granja}_${data.mes}.xlsx`
      : `FinalDeMes_${data.granja}_${data.mes}.xlsx`;

    const emailOpts = {
      from:    'Final de Mes <finaldemes@premierpigs.com>',
      to:      DESTINATARIOS,
      subject,
      html,
    };
    if (excelIndividual) {
      emailOpts.attachments = [
        { filename, content: excelIndividual.toString('base64') },
      ];
    }
    await resend.emails.send(emailOpts);
  } catch (err) {
    console.error('Error enviando email:', err);
    return res.status(500).json({ error: 'Error al enviar el email' });
  }

  // Comprobar si todas las granjas de madres han enviado este mes
  if (!esDestete) {
    try {
      const countResult = await sql`
        SELECT COUNT(DISTINCT granja) as total
        FROM inventario
        WHERE mes = ${data.mes} AND granja = ANY(${GRANJAS_MADRES_ACTIVAS})`;
      const totalGranjas = parseInt(countResult[0].total, 10);

      if (totalGranjas >= TOTAL_GRANJAS) {
        const madresRows = await sql`
          SELECT granja, data FROM inventario
          WHERE mes = ${data.mes} AND granja = ANY(${GRANJAS_MADRES})`;
        const desteteRows = await sql`
          SELECT granja, data FROM inventario
          WHERE mes = ${data.mes} AND granja = ANY(${GRANJAS_DESTETE})`;
        const mesLabel = getMonthLabel(data.mes);
        const excelBuffer = await generateExistencias(madresRows, desteteRows, data.mes);

        await resend.emails.send({
          from: 'Final de Mes <finaldemes@premierpigs.com>',
          to: DESTINATARIOS,
          subject: `EXISTENCIAS — Todas las granjas — ${mesLabel}`,
          html: `
            <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">
              <h2 style="color:#1a3a5c;">Existencias Completas</h2>
              <p style="color:#4a6080;">Todas las <strong>${totalGranjas} granjas</strong> han enviado el inventario de <strong>${mesLabel}</strong>.</p>
              <p style="color:#4a6080;">Adjunto el Excel de existencias con los totales sumados.</p>
              <hr style="border:none;border-top:1px solid #dde3ec;margin:20px 0;" />
              <p style="font-size:12px;color:#7a8fa8;">Enviado automáticamente · Premier Pigs</p>
            </div>`,
          attachments: [
            {
              filename: `EXISTENCIAS_${data.mes}.xlsx`,
              content: excelBuffer.toString('base64'),
            },
          ],
        });

        console.log(`Existencias enviadas: ${totalGranjas} granjas para ${data.mes}`);
      }
    } catch (summaryErr) {
      console.error('Error comprobando/enviando existencias:', summaryErr);
    }
  }

  return res.status(200).json({ ok: true });
}

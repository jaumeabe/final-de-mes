import { Resend } from 'resend';
import { neon } from '@neondatabase/serverless';
import ExcelJS from 'exceljs';

const resend = new Resend(process.env.RESEND_API_KEY);
const sql = neon(process.env.DATABASE_URL);

const DESTINATARIOS = [
  'jaumeabellana@gmail.com',
];

const EXISTENCIAS_FIELDS = [
  { key: 'num_machos', label: 'MACHOS' },
  { key: 'num_cerdas', label: 'CERDAS' },
  { key: 'num_primerizas', label: 'PRIMERIZAS' },
  { key: 'lechones_maternidad', label: 'LECHONES MATERNIDAD' },
  { key: 'lechones_destete', label: 'LECHONES DESTETE' },
];

function getMonthColumn(mesKey) {
  return parseInt(mesKey.split('-')[1], 10) + 1;
}

function getMonthLabel(mesKey) {
  const [y, m] = mesKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

function getPreviousMonthKey() {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

async function generateExistencias(allRows, mesKey) {
  const wb = new ExcelJS.Workbook();
  const year = mesKey.split('-')[0];
  const ws = wb.addWorksheet(year);
  const col = getMonthColumn(mesKey);

  const meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  const headerFont = { bold: true, size: 12, name: 'Arial', color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } };
  const labelFont = { bold: true, size: 11, name: 'Arial' };
  const dataFont = { size: 11, name: 'Arial' };
  const borderThin = {
    top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  };

  ws.getColumn(1).width = 24;
  for (let i = 2; i <= 13; i++) ws.getColumn(i).width = 14;

  // Fila 1: Título
  const titleRow = ws.getRow(1);
  titleRow.getCell(2).value = `EXISTENCIAS ${year}`;
  titleRow.getCell(2).font = { bold: true, size: 14, name: 'Arial', color: { argb: 'FF1A3A5C' } };
  ws.mergeCells('B1:M1');
  titleRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
  titleRow.height = 28;

  // Fila 2: Meses
  const monthRow = ws.getRow(2);
  for (let i = 0; i < 12; i++) {
    const cell = monthRow.getCell(i + 2);
    cell.value = meses[i];
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = borderThin;
  }
  monthRow.height = 22;

  // Sumar valores de todas las granjas
  const totals = {};
  for (const field of EXISTENCIAS_FIELDS) {
    totals[field.key] = 0;
    for (const row of allRows) {
      totals[field.key] += Number(row.data[field.key]) || 0;
    }
  }

  // Filas 3-7: Datos
  EXISTENCIAS_FIELDS.forEach((field, idx) => {
    const row = ws.getRow(3 + idx);
    row.getCell(1).value = field.label;
    row.getCell(1).font = labelFont;
    row.getCell(1).border = borderThin;

    const cell = row.getCell(col);
    cell.value = totals[field.key];
    cell.font = dataFont;
    cell.alignment = { horizontal: 'center' };
    cell.numFmt = '#,##0';

    for (let c = 2; c <= 13; c++) {
      row.getCell(c).border = borderThin;
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export default async function handler(req, res) {
  const authHeader = req.headers['authorization'];
  const isCron = authHeader === `Bearer ${process.env.CRON_SECRET}`;
  const isManual = req.query.key === process.env.CRON_SECRET;

  if (!isCron && !isManual) {
    return res.status(401).json({ error: 'No autorizado' });
  }

  const mesKey = req.query.mes || getPreviousMonthKey();

  let rows;
  try {
    rows = await sql`SELECT granja, data FROM inventario WHERE mes = ${mesKey}`;
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }

  if (!rows || rows.length === 0) {
    return res.status(200).json({ ok: true, message: 'No hay datos para este mes' });
  }

  const mesLabel = getMonthLabel(mesKey);
  const excelBuffer = await generateExistencias(rows, mesKey);

  try {
    await resend.emails.send({
      from: 'Final de Mes <onboarding@resend.dev>',
      to: DESTINATARIOS,
      subject: `EXISTENCIAS — ${mesLabel} (${rows.length} granjas)`,
      html: `
        <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">
          <h2 style="color:#1a3a5c;">Existencias — ${mesLabel}</h2>
          <p style="color:#4a6080;">Granjas reportadas: <strong>${rows.length}</strong></p>
          <p style="color:#4a6080;">Adjunto el Excel de existencias con los totales sumados de todas las granjas.</p>
          <hr style="border:none;border-top:1px solid #dde3ec;margin:20px 0;" />
          <p style="font-size:12px;color:#7a8fa8;">Enviado automáticamente · Premier Pigs</p>
        </div>`,
      attachments: [
        {
          filename: `EXISTENCIAS_${mesKey}.xlsx`,
          content: excelBuffer.toString('base64'),
        },
      ],
    });

    return res.status(200).json({ ok: true, message: `Existencias enviadas para ${mesLabel}` });
  } catch (err) {
    console.error('Error enviando existencias:', err);
    return res.status(500).json({ error: 'Error al enviar existencias' });
  }
}

import { Resend } from 'resend';
import { neon } from '@neondatabase/serverless';

const resend = new Resend(process.env.RESEND_API_KEY);
const sql = neon(process.env.DATABASE_URL);

const DESTINATARIOS = (process.env.EMAIL_TO || '').split(',').map(e => e.trim()).filter(Boolean);

const TODAS_GRANJAS = [
  'PORCELIBOR','CASTELLNOU','PI','SISALLAR 1','SENTERADA','SISALLAR 3',
  'CASERIO','MASIA','GRANADELLA','FAYON ABUELA','INDUSTRIAL','CARTUJA 2',
  'MARRUGAT','SINOGA','PASCUALET','NOVIPORCI','LES SERRES','ALFUSPI',
  'FUSTERO','GUINEU','SANTA ROSA','GIRALT','EL SOLER','VENDRELL',
  'PORDALL','RUBIO','ESCODA','CARTUJA 1','PORDECONA','SISALLAR 4',
];

function getPreviousMonthKey() {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

function getMonthLabel(mesKey) {
  const [y, m] = mesKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

export default async function handler(req, res) {
  const authHeader = req.headers['authorization'];
  const isCron = authHeader === `Bearer ${process.env.CRON_SECRET}`;
  const isManual = req.query.key === process.env.CRON_SECRET;

  if (!isCron && !isManual) {
    return res.status(401).json({ error: 'No autorizado' });
  }

  const mesKey = req.query.mes || getPreviousMonthKey();
  const mesLabel = getMonthLabel(mesKey);

  let rows;
  try {
    rows = await sql`SELECT DISTINCT granja FROM inventario WHERE mes = ${mesKey}`;
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }
  const reportadas = rows.map(r => r.granja);
  const pendientes = TODAS_GRANJAS.filter(g => !reportadas.includes(g));

  if (pendientes.length === 0) {
    return res.status(200).json({ ok: true, message: `Todas las granjas han reportado ${mesLabel}` });
  }

  const listaHtml = pendientes.map(g => `<li style="padding:4px 0;color:#1a3a5c;font-weight:600;">${g}</li>`).join('');

  try {
    await resend.emails.send({
      from: 'Final de Mes <onboarding@resend.dev>',
      to: DESTINATARIOS,
      subject: `⚠️ ${pendientes.length} granjas sin reportar — ${mesLabel}`,
      html: `
        <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">
          <h2 style="color:#b71c1c;">Granjas pendientes de reportar</h2>
          <p style="color:#4a6080;">Mes: <strong>${mesLabel}</strong></p>
          <p style="color:#4a6080;">Reportadas: <strong>${reportadas.length}</strong> de <strong>${TODAS_GRANJAS.length}</strong></p>
          <p style="color:#b71c1c;font-weight:600;">Faltan ${pendientes.length} granjas:</p>
          <ul style="list-style:none;padding:0;">${listaHtml}</ul>
          <hr style="border:none;border-top:1px solid #dde3ec;margin:20px 0;" />
          <p style="font-size:12px;color:#7a8fa8;">Aviso automático · Premier Pigs</p>
        </div>`,
    });

    return res.status(200).json({ ok: true, message: `Aviso enviado: ${pendientes.length} granjas pendientes` });
  } catch (err) {
    console.error('Error enviando aviso:', err);
    return res.status(500).json({ error: 'Error al enviar aviso' });
  }
}

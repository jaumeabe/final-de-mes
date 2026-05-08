import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

// ALFUSPI excluida temporalmente: pausada, sin animales hasta nuevo aviso
const TODAS_GRANJAS = [
  'PORCELIBOR','CASTELLNOU','PI','SISALLAR 1','SENTERADA','SISALLAR 3',
  'CASERIO','MASIA','GRANADELLA','FAYON ABUELA','INDUSTRIAL','CARTUJA 2',
  'MARRUGAT','SINOGA','PASCUALET','NOVIPORCI','LES SERRES',
  'FUSTERO','GUINEU','SANTA ROSA','GIRALT','EL SOLER','VENDRELL',
  'PORDALL','ESCODA','CARTUJA 1','PORDECONA','SISALLAR 4',
  'CEREALS','MAIALS','LLOBET','INGLES','ZAIDIN','JAUMANDREU','BORGES 1','RUBIO',
];

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

export default async function handler(req, res) {
  const authHeader = req.headers['authorization'];
  const isCron = authHeader === `Bearer ${process.env.CRON_SECRET}`;
  const isManual = req.query.key === process.env.CRON_SECRET || req.query.key === process.env.ADMIN_PASS;

  if (!isCron && !isManual) {
    return res.status(401).json({ error: 'No autorizado' });
  }

  const mesKey = req.query.mes || getPreviousMonthKey();

  if (!/^\d{4}-\d{2}$/.test(mesKey)) {
    return res.status(400).json({ error: 'Formato de mes no válido. Usa YYYY-MM' });
  }

  try {
    const rows = await sql`SELECT DISTINCT granja FROM inventario WHERE mes = ${mesKey}`;
    const reportadas = rows.map(r => r.granja);
    const pendientes = TODAS_GRANJAS.filter(g => !reportadas.includes(g));

    return res.status(200).json({
      mes: mesKey,
      mesLabel: getMonthLabel(mesKey),
      totalGranjas: TODAS_GRANJAS.length,
      reportadas: reportadas.length,
      pendientes: pendientes.length,
      granjasReportadas: reportadas.sort(),
      granjasPendientes: pendientes.sort(),
    });
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }
}

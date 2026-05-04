import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

// ALFUSPI excluida temporalmente: pausada, sin animales hasta nuevo aviso
const GRANJAS_MADRES = [
  'PORCELIBOR','CASTELLNOU','PI','SISALLAR 1','SENTERADA','SISALLAR 3',
  'CASERIO','MASIA','GRANADELLA','FAYON ABUELA','INDUSTRIAL','CARTUJA 2',
  'MARRUGAT','SINOGA','PASCUALET','NOVIPORCI','LES SERRES',
  'FUSTERO','GUINEU','SANTA ROSA','GIRALT','EL SOLER','VENDRELL',
  'PORDALL','ESCODA','CARTUJA 1','PORDECONA','SISALLAR 4',
];

const GRANJAS_DESTETE = [
  'CEREALS','MAIALS','LLOBET','INGLES','ZAIDIN','JAUMANDREU','BORGES 1','RUBIO',
];

const TODAS_GRANJAS = [...GRANJAS_MADRES, ...GRANJAS_DESTETE];

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
  res.setHeader('Access-Control-Allow-Origin', '*');

  const mesKey = req.query.mes || getPreviousMonthKey();

  if (!/^\d{4}-\d{2}$/.test(mesKey)) {
    return res.status(400).json({ error: 'Formato de mes no válido. Usa YYYY-MM' });
  }

  try {
    const rows = await sql`SELECT DISTINCT granja FROM inventario WHERE mes = ${mesKey}`;
    const reportadas = rows.map(r => r.granja);

    const madresPendientes = GRANJAS_MADRES.filter(g => !reportadas.includes(g));
    const desteePendientes = GRANJAS_DESTETE.filter(g => !reportadas.includes(g));
    const madresReportadas = GRANJAS_MADRES.filter(g => reportadas.includes(g));
    const desteteReportadas = GRANJAS_DESTETE.filter(g => reportadas.includes(g));

    return res.status(200).json({
      mes: mesKey,
      mesLabel: getMonthLabel(mesKey),
      total: TODAS_GRANJAS.length,
      totalReportadas: reportadas.length,
      totalPendientes: TODAS_GRANJAS.length - reportadas.length,
      madres: {
        total: GRANJAS_MADRES.length,
        reportadas: madresReportadas.sort(),
        pendientes: madresPendientes.sort(),
      },
      destete: {
        total: GRANJAS_DESTETE.length,
        reportadas: desteteReportadas.sort(),
        pendientes: desteePendientes.sort(),
      },
    });
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }
}

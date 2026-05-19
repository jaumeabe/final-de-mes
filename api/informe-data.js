import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

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

const CAMPOS_SUMA = [
  'num_machos','num_cerdas','num_primerizas','total_cerdas_primerizas',
  'lechones_maternidad','lechones_destete',
  'primerizas_entradas','machos_entrados',
  'venta_cerdas','venta_primerizas','venta_lechones_parideras',
  'venta_lechones_destete','venta_tostones',
  'num_partos','lechones_nacidos_vivos','cerdas_destetadas','lechones_destetados',
  'entrados_destete_propio','cubriciones_totales','cubriciones_conflictivo',
  'cubiertas_primera_vez','muerte_cerdas','muerte_primerizas',
  'bajas_destete','bajas_parideras',
];

function getPreviousMonthKey() {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function getMonthLabel(mesKey) {
  const [y, m] = mesKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[parseInt(m, 10) - 1]} ${y}`;
}

function avg(values) {
  const nums = values.filter(v => v != null && v !== '' && !isNaN(Number(v)));
  if (!nums.length) return null;
  return (nums.reduce((s, v) => s + Number(v), 0) / nums.length).toFixed(2);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const mesKey = req.query.mes || getPreviousMonthKey();

  if (!/^\d{4}-\d{2}$/.test(mesKey)) {
    return res.status(400).json({ error: 'Formato de mes no válido. Usa YYYY-MM' });
  }

  try {
    const madresRows = await sql`
      SELECT granja, data FROM inventario
      WHERE mes = ${mesKey} AND granja = ANY(${GRANJAS_MADRES})
      ORDER BY granja`;

    const desteteRows = await sql`
      SELECT granja, data FROM inventario
      WHERE mes = ${mesKey} AND granja = ANY(${GRANJAS_DESTETE})
      ORDER BY granja`;

    // Totales madres
    const totales = {};
    for (const campo of CAMPOS_SUMA) {
      totales[campo] = madresRows.reduce((s, r) => s + (Number(r.data[campo]) || 0), 0);
    }
    totales.prom_nacidos_vivos = avg(madresRows.map(r => r.data.prom_nacidos_vivos));
    totales.prom_destetados    = avg(madresRows.map(r => r.data.prom_destetados));

    // Totales destete
    const totalesDestete = {
      existencias: desteteRows.reduce((s, r) => s + (Number(r.data.existencias) || 0), 0),
      salidas:     desteteRows.reduce((s, r) => s + (Number(r.data.salidas)     || 0), 0),
      entradas:    desteteRows.reduce((s, r) => s + (Number(r.data.entradas)    || 0), 0),
      bajas:       desteteRows.reduce((s, r) => s + (Number(r.data.bajas)       || 0), 0),
      final_mes:   desteteRows.reduce((s, r) => s + (Number(r.data.final_mes)   || 0), 0),
    };

    return res.status(200).json({
      mes: mesKey,
      mesLabel: getMonthLabel(mesKey),
      madres: madresRows.map(r => ({ granja: r.granja, data: r.data })),
      destete: desteteRows.map(r => ({ granja: r.granja, data: r.data })),
      totalesMadres: totales,
      totalesDestete,
    });
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }
}

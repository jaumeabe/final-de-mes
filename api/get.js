import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { granja, mes } = req.query;

  if (!granja || !mes) {
    return res.status(400).json({ error: 'Faltan parámetros granja y mes' });
  }

  if (!/^\d{4}-\d{2}$/.test(mes)) {
    return res.status(400).json({ error: 'Formato de mes no válido' });
  }

  try {
    const rows = await sql`
      SELECT data, created_at
      FROM inventario
      WHERE granja = ${granja} AND mes = ${mes}
      LIMIT 1`;

    if (!rows || rows.length === 0) {
      return res.status(200).json({ found: false });
    }

    return res.status(200).json({
      found: true,
      data: rows[0].data,
      created_at: rows[0].created_at,
    });
  } catch (dbErr) {
    console.error('Error consultando DB:', dbErr);
    return res.status(500).json({ error: 'Error al consultar la base de datos' });
  }
}

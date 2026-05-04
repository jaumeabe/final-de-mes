import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  // Solo accesible con CRON_SECRET
  if (req.query.key !== process.env.CRON_SECRET) {
    return res.status(401).json({ error: 'No autorizado' });
  }

  try {
    // 1. Leer el registro erróneo (mayo 2026)
    const mayo = await sql`
      SELECT id, granja, mes, data
      FROM inventario
      WHERE granja = 'CARTUJA 2' AND mes = '2026-05'`;

    if (mayo.length === 0) {
      return res.status(200).json({ ok: false, message: 'No hay datos de CARTUJA 2 en 2026-05' });
    }

    const dataMayo = { ...mayo[0].data, mes: '2026-04' };

    // 2. Insertar/actualizar en abril 2026
    await sql`
      INSERT INTO inventario (granja, mes, data)
      VALUES ('CARTUJA 2', '2026-04', ${JSON.stringify(dataMayo)})
      ON CONFLICT (granja, mes)
      DO UPDATE SET data = ${JSON.stringify(dataMayo)}, created_at = NOW()`;

    // 3. Borrar el registro de mayo 2026
    await sql`DELETE FROM inventario WHERE granja = 'CARTUJA 2' AND mes = '2026-05'`;

    return res.status(200).json({
      ok: true,
      message: 'Datos de CARTUJA 2 movidos de 2026-05 a 2026-04 correctamente',
    });
  } catch (err) {
    console.error('Error:', err);
    return res.status(500).json({ error: err.message });
  }
}

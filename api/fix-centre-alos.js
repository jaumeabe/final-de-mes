import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  if (req.query.key !== process.env.CRON_SECRET) {
    return res.status(401).json({ error: 'No autorizado' });
  }

  const data = {
    granja: 'CENTRE ALOS',
    mes: '2026-04',
    tipo: 'machos',
    num_machos: 94,
  };

  try {
    await sql`
      INSERT INTO inventario (granja, mes, data)
      VALUES (${data.granja}, ${data.mes}, ${JSON.stringify(data)})
      ON CONFLICT (granja, mes)
      DO UPDATE SET data = ${JSON.stringify(data)}, created_at = NOW()`;

    return res.status(200).json({ ok: true, message: 'CENTRE ALOS 2026-04 guardado: 94 machos' });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}

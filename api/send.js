import { Resend } from 'resend';

const resend = new Resend(process.env.RESEND_API_KEY);

const DESTINATARIOS = [
  'timea@premierpigs.com',
    'jaume@premierpigs.com',
      'jaumejr@premierpigs.com',
      ];

      const LABELS = {
        num_machos:               'Número de machos',
          num_cerdas:               'Número de cerdas',
            num_primerizas:           'Número de primerizas',
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
                                                                cubriciones_totales:      'Cubriciones totales',
                                                                  cubriciones_conflictivo:  'Cubriciones origen conflictivo',
                                                                    cubiertas_primera_vez:    'Cubiertas por primera vez',
                                                                      bajas_destete:            'Bajas destete',
                                                                        bajas_parideras:          'Bajas parideras al mes',
                                                                        };

                                                                        const SECTIONS = [
                                                                          { title: 'INVENTARIO',       fields: ['num_machos','num_cerdas','num_primerizas','lechones_maternidad','lechones_destete'] },
                                                                            { title: 'ENTRADAS',         fields: ['primerizas_entradas','acumulado_primerizas','machos_entrados'] },
                                                                              { title: 'VENTAS',           fields: ['venta_cerdas','acum_venta_cerdas','venta_primerizas','acum_venta_primerizas','venta_lechones_parideras','acum_venta_parideras','venta_lechones_destete','acum_venta_destete','venta_tostones','acum_venta_tostones'] },
                                                                                { title: 'MUERTES',          fields: ['muerte_cerdas','acum_muerte_cerdas','muerte_primerizas','acum_muerte_primerizas'] },
                                                                                  { title: 'PARTOS Y DESTETE', fields: ['num_partos','lechones_nacidos_vivos','prom_nacidos_vivos','cerdas_destetadas','lechones_destetados','prom_destetados'] },
                                                                                    { title: 'CUBRICIONES',      fields: ['cubriciones_totales','cubriciones_conflictivo','cubiertas_primera_vez'] },
                                                                                      { title: 'BAJAS',            fields: ['bajas_destete','bajas_parideras'] },
                                                                                      ];

                                                                                      function buildHtml(data) {
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
                                                                                                                                                                                                                            
                                                                                                                                                                                                                            export default async function handler(req, res) {
                                                                                                                                                                                                                              if (req.method !== 'POST') {
                                                                                                                                                                                                                                  return res.status(405).json({ error: 'Method not allowed' });
                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                      const data = req.body;
                                                                                                                                                                                                                                      
                                                                                                                                                                                                                                        if (!data.granja || !data.mes) {
                                                                                                                                                                                                                                            return res.status(400).json({ error: 'Faltan campos obligatorios' });
                                                                                                                                                                                                                                              }
                                                                                                                                                                                                                                              
                                                                                                                                                                                                                                                try {
                                                                                                                                                                                                                                                    await resend.emails.send({
                                                                                                                                                                                                                                                          from:    'Final de Mes <onboarding@resend.dev>',
                                                                                                                                                                                                                                                                to:      DESTINATARIOS,
                                                                                                                                                                                                                                                                      subject: `Final de Mes — ${data.granja} — ${data.mes}`,
                                                                                                                                                                                                                                                                            html:    buildHtml(data),
                                                                                                                                                                                                                                                                                });
                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                    return res.status(200).json({ ok: true });
                                                                                                                                                                                                                                                                                      } catch (err) {
                                                                                                                                                                                                                                                                                          console.error('Error enviando email:', err);
                                                                                                                                                                                                                                                                                              return res.status(500).json({ error: 'Error al enviar el email' });
                                                                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                                                                }

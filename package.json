// api/sharepoint.js — Vercel Serverless Function
// Risolve il CORS e gestisce autenticazione + ricerca SharePoint

export default async function handler(req, res) {
  // CORS headers — permette chiamate da qualsiasi origine
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { action, query } = req.body || {};

  // Credenziali da variabili d'ambiente Vercel (sicure, non nel codice)
  const tenantId = process.env.SP_TENANT_ID;
  const clientId = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const siteUrl = process.env.SP_SITE_URL || 'https://carafasm.sharepoint.com/sites/Processi';

  if (!tenantId || !clientId || !clientSecret) {
    return res.status(500).json({ error: 'Credenziali server non configurate' });
  }

  try {
    // 1. Ottieni token da Microsoft
    const tokenResp = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      }
    );

    const tokenData = await tokenResp.json();
    if (!tokenData.access_token) {
      return res.status(401).json({ error: 'Autenticazione fallita', detail: tokenData.error_description });
    }

    const token = tokenData.access_token;
    const headers = { Authorization: `Bearer ${token}` };

    // 2. Trova il sito SharePoint
    const siteHostname = new URL(siteUrl).hostname;
    const sitePath = new URL(siteUrl).pathname;
    const siteResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
      { headers }
    );
    const siteData = await siteResp.json();

    if (!siteData.id) {
      return res.status(404).json({ error: 'Sito SharePoint non trovato', detail: siteData });
    }

    if (action === 'list') {
      // Lista tutti i file nella libreria documenti
      const drivesResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`,
        { headers }
      );
      const drivesData = await drivesResp.json();
      const drives = drivesData.value || [];

      let allFiles = [];

      for (const drive of drives) {
        // Recupera file ricorsivamente
        const filesResp = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/search(q='')`,
          { headers }
        );
        const filesData = await filesResp.json();
        const files = (filesData.value || [])
          .filter(f => f.file)
          .map(f => ({
            id: f.id,
            name: f.name,
            webUrl: f.webUrl,
            driveId: drive.id,
            folder: f.parentReference?.path?.split('root:')[1] || '/',
          }));
        allFiles = [...allFiles, ...files];
      }

      return res.status(200).json({ files: allFiles });
    }

    if (action === 'search' && query) {
      // Cerca documenti pertinenti alla query
      const searchResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/search(q='${encodeURIComponent(query)}')`,
        { headers }
      );
      const searchData = await searchResp.json();
      const results = (searchData.value || [])
        .filter(f => f.file)
        .slice(0, 5)
        .map(f => ({
          name: f.name,
          webUrl: f.webUrl,
          folder: f.parentReference?.path?.split('root:')[1] || '/',
        }));

      return res.status(200).json({ results });
    }

    return res.status(400).json({ error: 'Azione non valida. Usa action: list o search' });

  } catch (err) {
    return res.status(500).json({ error: 'Errore server', detail: err.message });
  }
}

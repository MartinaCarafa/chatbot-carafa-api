export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { action, query } = req.body || {};
  const tenantId = process.env.SP_TENANT_ID;
  const clientId = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const siteUrl = process.env.SP_SITE_URL || 'https://carafam.sharepoint.com/sites/Processi';

  if (!tenantId || !clientId || !clientSecret) {
    return res.status(500).json({ error: 'Credenziali server non configurate' });
  }

  try {
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
      return res.status(401).json({ error: 'Auth fallita', detail: tokenData.error_description });
    }

    const token = tokenData.access_token;
    const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };

    const siteHostname = new URL(siteUrl).hostname;
    const sitePath = new URL(siteUrl).pathname;
    const siteResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
      { headers }
    );
    const siteData = await siteResp.json();

    if (!siteData.id) {
      return res.status(404).json({ error: 'Sito non trovato', url: siteUrl, detail: siteData });
    }

    if (action === 'search' && query) {
      const searchResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/search(q='${encodeURIComponent(query)}')`,
        { headers }
      );
      const searchData = await searchResp.json();
      const results = (searchData.value || [])
        .filter(f => f.file)
        .slice(0, 5)
        .map(f => ({ name: f.name, webUrl: f.webUrl }));
      return res.status(200).json({ results });
    }

    if (action === 'list') {
      const drivesResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`,
        { headers }
      );
      const drivesData = await drivesResp.json();
      let allFiles = [];
      for (const drive of (drivesData.value || [])) {
        const filesResp = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/children`,
          { headers }
        );
        const filesData = await filesResp.json();
        const files = (filesData.value || []).filter(f => f.file).map(f => ({
          name: f.name, webUrl: f.webUrl
        }));
        allFiles = [...allFiles, ...files];
      }
      return res.status(200).json({ files: allFiles });
    }

    return res.status(400).json({ error: 'Azione non valida' });

  } catch (err) {
    return res.status(500).json({ error: 'Errore server', detail: err.message });
  }
}

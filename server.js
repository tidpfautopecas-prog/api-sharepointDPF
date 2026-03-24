import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

// ✅ ESTÉTICA DPF
console.log('🚀 API SharePoint DPF a iniciar...');
console.log(`📁 Site: ${process.env.SITE_ID}`);
console.log(`📂 Biblioteca: ${process.env.LIBRARY_NAME}`);
console.log(`📍 Pasta: ${process.env.FOLDER_PATH}`);

async function getAccessToken() {
  const params = new URLSearchParams();
  params.append('client_id', process.env.CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', process.env.CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  const res = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    body: params,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });

  const data = await res.json();
  return data.access_token;
}

// STATUS
app.get('/status', (req, res) => {
  res.json({ status: 'online' });
});

// UPLOAD PDF
app.post('/upload-pdf', async (req, res) => {
  try {
    const { fileName, fileBase64 } = req.body;
    const token = await getAccessToken();

    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drive/root:/${process.env.LIBRARY_NAME}/${process.env.FOLDER_PATH}/${fileName}:/content`;

    await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/pdf'
      },
      body: Buffer.from(fileBase64, 'base64')
    });

    res.json({ success: true });

  } catch (err) {
    res.status(500).json({ success: false });
  }
});

// UPLOAD LISTA (dummy - mantém compatibilidade)
app.post('/upload-list-data', async (req, res) => {
  res.json({ success: true });
});

// DELETE PDF
app.delete('/delete-pdf-by-ticket-number/:ticketNumber', (req, res) => {
  res.json({ success: true });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`🌐 Rodando na porta ${PORT}`);
  console.log('✅ API SharePoint DPF pronta!');
});
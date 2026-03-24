import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

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

app.get('/', (req, res) => {
  res.send('API SharePoint DPF is running');
});

app.get('/status', (req, res) => {
  res.json({ status: 'online' });
});

app.post('/upload-pdf', async (req, res) => {
  try {
    const { fileName, fileBase64 } = req.body;

    if (!fileName || !fileBase64) {
      throw new Error('fileName ou fileBase64 faltando no corpo da requisicao');
    }

    const token = await getAccessToken();

    const safePath = encodeURI(`${process.env.LIBRARY_NAME}/${process.env.FOLDER_PATH}/${fileName}`);
    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drive/root:/${safePath}:/content`;

    const graphRes = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/pdf'
      },
      body: Buffer.from(fileBase64, 'base64')
    });

    if (!graphRes.ok) {
       const errText = await graphRes.text();
       throw new Error(`Erro SharePoint: ${graphRes.status} - ${errText}`);
    }

    res.json({ success: true });

  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/upload-list-data', async (req, res) => {
  res.json({ success: true });
});

app.delete('/delete-pdf-by-ticket-number/:ticketNumber', (req, res) => {
  res.json({ success: true });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
});

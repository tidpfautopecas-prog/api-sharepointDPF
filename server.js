import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

// ================= TOKEN =================
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

// ================= PEGAR ID DA LISTA =================
async function getListIdByName(token, listName) {
  const listsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`;

  const res = await fetch(listsUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const data = await res.json();

  const list = data.value?.find(
    (l) => l.displayName === listName || l.name === listName
  );

  if (!list) {
    throw new Error(`Lista '${listName}' não encontrada`);
  }

  return list.id;
}

// ================= ROTAS =================
app.get('/', (req, res) => {
  res.send('API SharePoint DPF is running');
});

app.get('/status', (req, res) => {
  res.json({ status: 'online' });
});

// ================= DEBUG =================
app.get('/debug-sharepoint', async (req, res) => {
  try {
    const token = await getAccessToken();

    const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    const listsData = await listsRes.json();

    const laudoList = listsData.value.find(
      (l) => l.displayName === process.env.LIST_NAME || l.name === process.env.LIST_NAME
    );

    const colsRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${laudoList.id}/columns`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const colsData = await colsRes.json();

    res.json({
      listId: laudoList.id,
      colunas: colsData.value.map(c => ({
        NomeNaTela: c.displayName,
        NomeInterno: c.name
      }))
    });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ================= UPLOAD PDF =================
app.post('/upload-pdf', async (req, res) => {
  try {
    const { fileName, fileBase64 } = req.body;

    const token = await getAccessToken();

    const safePath = encodeURI(`${process.env.LIBRARY_NAME}/${process.env.FOLDER_PATH}/${fileName}`);

    const graphRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drive/root:/${safePath}:/content`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/pdf'
        },
        body: Buffer.from(fileBase64, 'base64')
      }
    );

    if (!graphRes.ok) {
      throw new Error(await graphRes.text());
    }

    res.json({ success: true });

  } catch (err) {
    console.error("ERRO PDF:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ================= UPLOAD LISTA =================
app.post('/upload-list-data', async (req, res) => {
  try {
    const { listData } = req.body;

    if (!listData?.length) {
      return res.json({ success: true });
    }

    const token = await getAccessToken();
    const listId = await getListIdByName(token, process.env.LIST_NAME);

    for (const item of listData) {

      // 🔥 USANDO NOMES VISÍVEIS (resolve erro 500)
      const fields = {
        "Título": item.Titulo || '',
        "Nº do ticket": item.ticketNumber || '',
        "Nome do Cliente": item.nomeCliente || '',
        "Item": item.item || '',
        "Qtde": item.qtde || '',
        "Motivo": item.motivo || '',
        "Origem do defeito": item.origemDefeito || '',
        "Disposição": item.disposicao || '',
        "Disposição das peças": item.disposicaoPecas || '',
        "Data de Geração": item.dataGeracao || ''
      };

      // 🔥 FOTOS
      const fotosMap = {
        foto1: "Foto 1",
        foto2: "Foto 2",
        foto3: "Foto 3",
        foto4: "Foto 4",
        foto5: "Foto 5",
        foto6: "Foto 6",
        foto7: "Foto 7",
        foto8: "Foto 8",
        foto9: "Foto 9",
        foto10: "Foto 10"
      };

      Object.keys(fotosMap).forEach(key => {
        if (item[key]) {
          fields[fotosMap[key]] = item[key];
        }
      });

      const graphRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ fields })
        }
      );

      if (!graphRes.ok) {
        const errText = await graphRes.text();
        throw new Error(`Erro SharePoint Lista: ${errText}`);
      }
    }

    res.json({ success: true });

  } catch (err) {
    console.error("ERRO LISTA:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ================= DELETE =================
app.delete('/delete-pdf-by-ticket-number/:ticketNumber', (req, res) => {
  res.json({ success: true });
});

// ================= START =================
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});

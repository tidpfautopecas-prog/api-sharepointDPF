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
  try {
    const { listData } = req.body;

    if (!listData || !Array.isArray(listData) || listData.length === 0) {
      return res.json({ success: true });
    }

    const token = await getAccessToken();
    const listName = process.env.LIST_NAME;

    for (const item of listData) {
      const fields = {
        Title: item.Title || '',
        N_x00b0__x0020_do_x0020_ticket: item.ticketNumber || '',
        Nome_x0020_do_x0020_Cliente: item.nomeCliente || '',
        Item: item.item || '',
        Qtde: item.qtde || '',
        Motivo: item.motivo || '',
        Origem_x0020_do_x0020_defeito: item.origemDefeito || '',
        Disposi_x00e7__x00e3_o: item.disposicao || '',
        Disposi_x00e7__x00e3_o_x0020_das_x0020_pe_x00e7_as: item.disposicaoPecas || '',
        Data_x0020_de_x0020_Gera_x00e7__x00e3_o: item.dataGeracao || ''
      };

      for (let i = 1; i <= 10; i++) {
        if (item[`foto${i}`]) {
          fields[`Foto_x0020_${i}`] = item[`foto${i}`];
        }
      }

      const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listName}/items`;

      const graphRes = await fetch(graphUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ fields })
      });

      if (!graphRes.ok) {
        const errText = await graphRes.text();
        throw new Error(`Erro SharePoint Lista: ${graphRes.status} - ${errText}`);
      }
    }

    res.json({ success: true });

  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/debug-sharepoint', async (req, res) => {
  try {
    const token = await getAccessToken();

    const listsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`;
    const listsRes = await fetch(listsUrl, { headers: { Authorization: `Bearer ${token}` } });
    const listsData = await listsRes.json();

    if (!listsData.value) {
        return res.json({ error: "Não foi possível carregar as listas", detalhes: listsData });
    }

    const laudoList = listsData.value.find(l => l.displayName === 'Laudo' || l.name === 'Laudo');

    if (!laudoList) {
       return res.json({ 
         aviso: "Lista 'Laudo' não encontrada com esse nome exato.", 
         listasDisponiveis: listsData.value.map(l => l.displayName) 
       });
    }

    const colsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${laudoList.id}/columns`;
    const colsRes = await fetch(colsUrl, { headers: { Authorization: `Bearer ${token}` } });
    const colsData = await colsRes.json();

    const colunasMapeadas = colsData.value.map(c => ({
        NomeNaTela: c.displayName,
        NomeInternoParaAPI: c.name
    }));

    res.json({
      sucesso: true,
      listId: laudoList.id,
      colunas: colunasMapeadas
    });

  } catch (err) {
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

app.delete('/delete-pdf-by-ticket-number/:ticketNumber', (req, res) => {
  res.json({ success: true });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
});

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

// 🔥 FUNÇÃO PARA PEGAR ID REAL DA LISTA
async function getListIdByName(token: string, listName: string) {
  const listsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`;

  const res = await fetch(listsUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const data = await res.json();

  const list = data.value?.find(
    (l: any) => l.displayName === listName || l.name === listName
  );

  if (!list) {
    throw new Error(`Lista '${listName}' não encontrada`);
  }

  return list.id;
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
      throw new Error('fileName ou fileBase64 faltando');
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

  } catch (err: any) {
    console.error("ERRO PDF:", err);
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

    // 🔥 CORREÇÃO: pegar ID da lista
    const listId = await getListIdByName(token, process.env.LIST_NAME);

    for (const item of listData) {

      // 🔥 CAMPOS CORRIGIDOS
      const fields = {
        Title: item.Title || '',
        N_x00ba__x0020_do_x0020_ticket: item.ticketNumber || '', // ✔️ CORRIGIDO (º)
        Nome_x0020_do_x0020_Cliente: item.nomeCliente || '',
        Item: item.item || '',
        Qtde: item.qtde || '',
        Motivo: item.motivo || '',
        Origem_x0020_do_x0020_defeito: item.origemDefeito || '',
        Disposi_x00e7__x00e3_o: item.disposicao || '',
        Disposi_x00e7__x00e3_o_x0020_das_x0020_pe_x00e7_as: item.disposicaoPecas || '',
        Data_x0020_de_x0020_Gera_x00e7__x00e3_o: item.dataGeracao || ''
      };

      // Fotos
      for (let i = 1; i <= 10; i++) {
        if (item[`foto${i}`]) {
          fields[`Foto_x0020_${i}`] = item[`foto${i}`];
        }
      }

      const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`;

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

  } catch (err: any) {
    console.error("ERRO LISTA:", err); // 🔥 LOG REAL
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/debug-sharepoint', async (req, res) => {
  try {
    const token = await getAccessToken();

    const listsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`;
    const listsRes = await fetch(listsUrl, { headers: { Authorization: `Bearer ${token}` } });
    const listsData = await listsRes.json();

    const laudoList = listsData.value.find(
      (l: any) => l.displayName === process.env.LIST_NAME || l.name === process.env.LIST_NAME
    );

    if (!laudoList) {
      return res.json({
        erro: "Lista não encontrada",
        listas: listsData.value.map((l: any) => l.displayName)
      });
    }

    const colsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${laudoList.id}/columns`;
    const colsRes = await fetch(colsUrl, { headers: { Authorization: `Bearer ${token}` } });
    const colsData = await colsRes.json();

    res.json({
      listId: laudoList.id,
      colunas: colsData.value.map((c: any) => ({
        NomeNaTela: c.displayName,
        NomeInterno: c.name
      }))
    });

  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/delete-pdf-by-ticket-number/:ticketNumber', (req, res) => {
  res.json({ success: true });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});

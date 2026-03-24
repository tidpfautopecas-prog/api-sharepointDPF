import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin'],
    credentials: true
}));

app.options('*', cors()); 
app.use(bodyParser.json({ limit: '50mb' }));

console.log('🚀 API SharePoint DPF iniciando...');

// 🔥 USANDO NOMES VISÍVEIS (FUNCIONA EM QUALQUER SHAREPOINT)
const COLUMN_MAPPING = {
    'Title': (row) => `${row.ticketNumber} - ${row.item} - ${row.motivo}`,
    'Nº do ticket': (row) => row.ticketNumber,
    'Nome do Cliente': (row) => row.nomeCliente,
    'Item': (row) => row.item,
    'Qtde': (row) => String(row.qtde),
    'Motivo': (row) => row.motivo,
    'Origem do defeito': (row) => row.origemDefeito,
    'Disposição': (row) => row.disposicao,
    'Disposição das peças': (row) => row.disposicaoPecas,
    'Data de Geração': (row) => row.dataGeracao || '',
    'Foto 1': (row) => row.foto1 || null,
    'Foto 2': (row) => row.foto2 || null,
    'Foto 3': (row) => row.foto3 || null,
    'Foto 4': (row) => row.foto4 || null,
    'Foto 5': (row) => row.foto5 || null,
    'Foto 6': (row) => row.foto6 || null,
    'Foto 7': (row) => row.foto7 || null,
    'Foto 8': (row) => row.foto8 || null,
    'Foto 9': (row) => row.foto9 || null,
    'Foto 10': (row) => row.foto10 || null,
};

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
    if (!data.access_token) throw new Error('Erro token');
    return data.access_token;
}

async function getListId(token) {
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`, {
        headers: { Authorization: `Bearer ${token}` }
    });

    const data = await res.json();

    const list = data.value.find(l => 
        l.displayName === process.env.LIST_NAME || l.name === process.env.LIST_NAME
    );

    if (!list) throw new Error('Lista não encontrada');

    return list.id;
}

// ================= ROTAS =================

app.get('/', (req, res) => {
    res.json({ status: 'online' });
});

// ================= PDF =================

app.post('/upload-pdf', async (req, res) => {
    try {
        const { fileName, fileBase64 } = req.body;

        const token = await getAccessToken();

        const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drive/root:/${process.env.LIBRARY_NAME}/${process.env.FOLDER_PATH}/${fileName}:/content`;

        const response = await fetch(url, {
            method: 'PUT',
            headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/pdf'
            },
            body: Buffer.from(fileBase64, 'base64')
        });

        if (!response.ok) throw new Error(await response.text());

        res.json({ success: true });

    } catch (err) {
        console.error('❌ ERRO PDF:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// ================= LISTA =================

app.post('/upload-list-data', async (req, res) => {
    try {
        const { listData } = req.body;

        if (!listData?.length) return res.json({ success: true });

        const token = await getAccessToken();
        const listId = await getListId(token);

        for (const row of listData) {

            const fields = {};

            for (const key in COLUMN_MAPPING) {
                const value = COLUMN_MAPPING[key](row);
                if (value !== null && value !== undefined && value !== '') {
                    fields[key] = value;
                }
            }

            const response = await fetch(
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

            if (!response.ok) {
                const err = await response.text();
                throw new Error(err);
            }
        }

        res.json({ success: true });

    } catch (err) {
        console.error('❌ ERRO LISTA:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// ================= START =================

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🌐 API rodando na porta ${PORT}`));

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
    console.error("Erro no upload da lista:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

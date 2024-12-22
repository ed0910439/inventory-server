const axios = require('axios');
const cheerio = require('cheerio');
const { app, Archive } = require('./server');

// �s�W API ���I�H�ҰʽL�I�y�{�M�ާ@���[�ƾ�
app.get('/api/startInventory', async (req, res) => {
    try {
        // 1: ?��?�e����M�۫ض��X�W?
        const currentDate = new Date();
        const year = currentDate.getFullYear();
        const month = String(currentDate.getMonth() + 1).padStart(2, '0');
        const collectionName = `${year}${month}��L`; // ??�۫ض��X�W?


        // 2: ? archives ?���̪�?�v?�u
        const latestInventory = await Archive.find().sort({ �L�I���: -1 }).limit(1).exec();

        // ?�v?�u??
        let historicalData = [];
        if (latestInventory.length > 0) {
            historicalData = latestInventory.map(item => ({
                �ӫ~�s��: item.�ӫ~�s��,
                ���: item.���,
            }));
        }

        // 3: ?�N�z��?����s?�u (HTML �榡)
        const response = await axios.get('https://epos.kingza.com.tw:8090/hyisoft.lost/exportpand.aspx?t=panDianItemCS&id=3148&ClassStore_fCheckSetID=');
        const html = response.data;
        const $ = cheerio.load(html);

        // �ѪR HTML ���?�e
        const newProducts = [];
        $('table tr').each((i, el) => {
            if (i === 0) return; // �����Ĥ@���?

            const row = $(el).find('td').map((j, cell) => $(cell).text().trim()).get();

            if (row.length >= 4) { // �̫O�ܤ֦�4�C
                const product = {
                    �ӫ~�s��: row[0] || '����',
                    �ӫ~�W��: row[1] || '',
                    �W��: '',
                    ���: row[3] || '���]�w',
                    �t��: '',
                    �w�O: '',
                };

                // ?�s�J���Ī�?�~
                if (product.�ӫ~�s�� && product.���) {
                    newProducts.push(product);
                }
            }
        });

        // 4: ??�X�s�~�A�קK�w���~?
        const existingProductIds = new Set(historicalData.map(product => product.�ӫ~�s��));
        const filteredNewProducts = newProducts.filter(product => !existingProductIds.has(product.�ӫ~�s��));

        // ��^�s�~�H��
        res.json(filteredNewProducts);

    } catch (error) {
        console.error('?�z?�l???�D?�X?:', error);
        res.status(500).send('�A?��??');
    }
});

const axios = require('axios');
const cheerio = require('cheerio');
const { app, Archive } = require('./server');

// 新增 API 端點以啟動盤點流程和操作附加數據
app.get('/api/startInventory', async (req, res) => {
    try {
        // 1: ?取?前日期和构建集合名?
        const currentDate = new Date();
        const year = currentDate.getFullYear();
        const month = String(currentDate.getMonth() + 1).padStart(2, '0');
        const collectionName = `${year}${month}月盤`; // ??构建集合名?


        // 2: ? archives ?取最近的?史?据
        const latestInventory = await Archive.find().sort({ 盤點日期: -1 }).limit(1).exec();

        // ?史?据??
        let historicalData = [];
        if (latestInventory.length > 0) {
            historicalData = latestInventory.map(item => ({
                商品編號: item.商品編號,
                單位: item.單位,
            }));
        }

        // 3: ?代理端?抓取新?据 (HTML 格式)
        const response = await axios.get('https://epos.kingza.com.tw:8090/hyisoft.lost/exportpand.aspx?t=panDianItemCS&id=3148&ClassStore_fCheckSetID=');
        const html = response.data;
        const $ = cheerio.load(html);

        // 解析 HTML 表格?容
        const newProducts = [];
        $('table tr').each((i, el) => {
            if (i === 0) return; // 忽略第一行表?

            const row = $(el).find('td').map((j, cell) => $(cell).text().trim()).get();

            if (row.length >= 4) { // 确保至少有4列
                const product = {
                    商品編號: row[0] || '未知',
                    商品名稱: row[1] || '',
                    規格: '',
                    單位: row[3] || '未設定',
                    廠商: '',
                    庫別: '',
                };

                // ?存入有效的?品
                if (product.商品編號 && product.單位) {
                    newProducts.push(product);
                }
            }
        });

        // 4: ??出新品，避免已有品?
        const existingProductIds = new Set(historicalData.map(product => product.商品編號));
        const filteredNewProducts = newProducts.filter(product => !existingProductIds.has(product.商品編號));

        // 返回新品信息
        res.json(filteredNewProducts);

    } catch (error) {
        console.error('?理?始???求?出?:', error);
        res.status(500).send('服?器??');
    }
});

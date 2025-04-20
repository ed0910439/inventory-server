const express = require('express');
const { Server } = require('socket.io');
const cookieParser = require('cookie-parser');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const http = require('http');
const ExcelJS = require('exceljs');
const axios = require('axios');
const cheerio = require('cheerio');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const morgan = require('morgan'); // 引入 morgan
//const csrf = require('csurf');
const bodyParser = require('body-parser');
//const morgan = require('morgan'); // 新增日誌中介

require('dotenv').config();

const app = express();

// 中介配置
//app.use(morgan('combined')); // 使用 morgan 記錄 HTTP 請求
app.use(cookieParser());
app.use(cors());
app.use(bodyParser.json({ limit: '10mb' })); // 更新這裡
app.use(bodyParser.urlencoded({ limit: '10mb', extended: true }));
app.use(helmet());
app.enable('trust proxy');
app.set('trust proxy', 1); // 1 是 'X-Forwarded-For' 的第一層代理
app.use(morgan('tiny')); 

/*// 設定 CSRF 保護
const csrfProtection = csrf({ cookie: true });
app.use(csrfProtection);

// 提供 CSRF 令牌的 API 端點
app.get('/api/csrf-token', (req, res) => {
 res.json({ csrfToken: req.csrfToken() });
});

// CSRF 錯誤處理
app.use((err, req, res, next) => {
 if (err.code === 'EBADCSRFTOKEN') {
 return res.status(403).json({ error: 'CSRF token validation failed' });
 }
 // 處理其他錯誤
 return res.status(500).json({ error: 'Something went wrong' });
});*/
app.use(cors({ origin: '*' })); // 或使用 '*' 來允許所有來源

// 配置 API 請求的速率限制，防止濫用
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 分鐘窗口
    max: 100, // 每個 IP 15 分鐘內最多可以請求 100 次
    message: '您發送請求的速度太快，麻煩您過五分鐘後再試！ '
});
app.use('/api/', limiter); // 只對 API 請求應用 rate limit



// 連接到 MongoDB
mongoose.connect(process.env.MONGODB_URI, {
    ssl: true,
})
    .then(() => console.log('成功連接到 MongoDB'))
    .catch(err => console.error('MongoDB 連接錯誤:', err));

// 定義產品模型
// 初始化 Express 應用後
const productSchema = new mongoose.Schema({
    停用: { type: Boolean, required: true },
    商品編號: { type: String, required: true },
    商品名稱: { type: String, required: false },
    規格: { type: String, required: false },
    數量: { type: Number, rquired: false },
    單位: { type: String, required: false },
    到期日: { type: String, required: false },
    廠商: { type: String, required: false },
    庫別: { type: String, required: false },
    盤點日期: { type: String, required: false },
    期初庫存: { type: String, required: false },

});

// 定義 sanitizeInput 函數
const sanitizeInput = (input) => {
    return encodeURIComponent(input.trim());
};
// 動態產生集合名稱
const currentDate = new Date();
let year = currentDate.getFullYear();
let lastYear = currentDate.getFullYear();
let month, lastMonth;

// 取得當前日期的日
let day = currentDate.getDate();

// 根據日期決定月份
if (day <= 15) {
    // 每月15日（含）以前
    month = currentDate.getMonth(); // 上個月
    lastMonth = currentDate.getMonth() - 1; // 上一個月

    // 若上一個月為 -1(即1月)，則需要調整年份
    if (month < 0) {
        month = 11; // 12月
        year -= 1; // 前一年
    }

    // 如果上上一個月小於0則也需要調整
    if (lastMonth < 0) {
        lastMonth = 11; // 12月
        lastYear -= 1; // 在調整
    }
} else {
    // 每月16日開始
    month = currentDate.getMonth() + 1; // 當前月份（1-12）
    lastMonth = currentDate.getMonth(); // 上個月
}

// 格式化月份為兩位數
const urlFormattedMonth = String(currentDate.getMonth() +1 ).padStart(2, '0');
const formattedMonth = String(month).padStart(2, '0');
const formattedLastMonth = String(lastMonth).padStart(2, '0'); // 轉換成1-12格式

console.log(year, formattedMonth, day); // 輸出當前年份、月份、日期
console.log(lastYear, formattedLastMonth, day); // 輸出上月份的年份、月份、日期


app.get('/api/testInventoryTemplate/:storeName', (req, res) => {
    const storeName = req.params.storeName;

    // 返回用于测试的示例数据
    const mockInventoryData = [
        { 停用: 'ture', 商品編號: '001', 商品名稱: '項目A', 規格: '', 廠商: '待設定', 庫別: '待設定' },
        { 停用: 'false', 商品編號: '002', 商品名稱: '項目B', 規格: '', 廠商: '待設定', 庫別: '待設定' },
        // 添加更多的模拟数据
    ];

    res.json(mockInventoryData);
});

app.get('/api/startInventory/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName

    try {
        if (storeName === 'notStart') {
            return res.status(204).send('尚未選擇門市'); // 使用 204
        } else {
            const today = `${year}-${urlFormattedMonth}-${day}`;
            const collectionName = `${year}${formattedMonth}${storeName}_tmp`; // 根據年份、月份和門市產生暫存集合名稱
            const newCollectionName = `${year}${formattedMonth}${storeName}`; // 當月正式集合名稱
            const lastCollectionName = `${lastYear}${formattedLastMonth}${storeName}`; // 上個月集合名稱
            const Product = mongoose.model(collectionName, productSchema); // 暫存模型
            const NewProduct = mongoose.model(newCollectionName, productSchema); // 當月正式模型
            const firstUrl = process.env.FIRST_URL.replace('${today}', today); // 替換 URL 中的變數
            const secondUrl = process.env.SECOND_URL;
            console.log('第一份 HTML URL:', firstUrl);
            console.log('第二份 HTML URL:', secondUrl);

            await Product.deleteMany(); // 清空暫存集合
            console.log(`已清空暫存集合: ${collectionName}`);
            console.log('當月暫存集合名稱:', collectionName);
            console.log('上個月集合名稱:', lastCollectionName);
            console.log('開始抓取 HTML 資料...');

            const firstResponse = await axios.get(firstUrl);
            const firstHtml = firstResponse.data;
            const $first = cheerio.load(firstHtml);

            const newProductsFromHtml = [];
            $first('table tr').each((i, el) => {
                if (i === 0) return; // 忽略表頭
                const row = $first(el).find('td').map((j, cell) => $first(cell).text().trim()).get();

                if (row.length > 3) {
                    const product = {
                        模板名稱: row[1],
                        商品編號: row[9],
                        商品名稱: row[10],
                        規格: row[11],
                    };
                    if (product.模板名稱 === '段純貞' && product.商品編號) { // 確保有商品編號
                        newProductsFromHtml.push(product); // 只保存有效的段純貞產品
                    }
                }
            });
            console.log(`從第一個資料來源抓取到 ${newProductsFromHtml.length} 個段純貞產品`);

            // 取得上個月的盤點資料並建立商品編號索引
            let lastMonthInventory = [];
            const lastMonthInventoryMap = new Map();
            try {
                const sourceCollection = mongoose.connection.collection(lastCollectionName);
                lastMonthInventory = await sourceCollection.find({}).toArray(); // 取得上個月集合所有資料
                lastMonthInventory.forEach(item => lastMonthInventoryMap.set(item.商品編號, item));
                console.log(`成功取得上個月盤點資料，共 ${lastMonthInventory.length} 筆`);
            } catch (error) {
                console.warn(`無法找到上個月的盤點集合: ${lastCollectionName}，將視為首次盤點。`);
                lastMonthInventory = []; // 如果上個月集合不存在，視為空陣列
            }

            // 建立本月盤點所需的基礎資料
            const refinedData = newProductsFromHtml.map(htmlProduct => {
                const lastMonthData = lastMonthInventoryMap.get(htmlProduct.商品編號);
                const is停用InHtml = htmlProduct.商品名稱.includes('停用') || htmlProduct.商品名稱.includes('勿下');

                return {
                    停用: lastMonthData?.停用 !== undefined ? lastMonthData.停用 : is停用InHtml,
                    商品編號: htmlProduct.商品編號,
                    商品名稱: htmlProduct.商品名稱,
                    規格: lastMonthData?.規格 || '',
                    數量: '',
                    單位: '', // 暫不處理
                    到期日: '',
                    廠商: lastMonthData?.廠商 || (htmlProduct.商品編號.includes('KO') || htmlProduct.商品編號.includes('KL') ? '王座(用)' : htmlProduct.商品編號.includes('KM') ? '央廚' : '待設定'),
                    庫別: lastMonthData?.庫別 || (lastMonthData?.停用 ? '未使用' : is停用InHtml ? '未使用' : '待設定'),
                    盤點日期: '',
                    期初庫存: lastMonthData?.數量 || '無數據',
                };
            });

            if (refinedData.length > 0) {
                await Product.insertMany(refinedData); // 存入暫存集合
                console.log(`已將 ${refinedData.length} 筆基礎產品資訊存入暫存集合: ${collectionName}`);
            } else {
                console.log('沒有需要新增的產品。');
            }

            // 建立暫存集合的商品編號映射
            const tempProductCodeMap = new Map(refinedData.map(item => [item.商品編號, item]));

            // 從第二個 HTML 資料來源抓取資料並更新單位
            console.log(`抓取第二個 HTML 資料...`);
            const secondResponse = await axios.get(secondUrl);
            const secondHtml = secondResponse.data;
            const $second = cheerio.load(secondHtml);

            const secondInventoryData = [];
            $second('table tr').each((i, el) => {
                if (i === 0) return; // 忽略表頭
                const row = $second(el).find('td').map((j, cell) => $second(cell).text().trim()).get();

                if (row.length > 3 && row[0] && row[3]) { // 確保商品編號和單位存在
                    secondInventoryData.push({
                        商品編號: row[0],
                        單位: row[3],
                    });
                }
            });
            console.log(`從第二個資料來源抓取到 ${secondInventoryData.length} 筆產品單位資訊`);

            // 更新暫存集合中的產品單位
            secondInventoryData.forEach(item => {
                const tempProduct = tempProductCodeMap.get(item.商品編號);
                if (tempProduct) {
                    tempProduct.單位 = item.單位;
                }
            });

            const updatedProducts = Array.from(tempProductCodeMap.values());

            // 傳回所有庫別為「待設定」新品項，等待使用者填寫
            const pendingProducts = updatedProducts.filter(product => product.庫別 === '待設定');

            if (pendingProducts.length === 0) {
                // 沒有待設定產品，將暫存資料轉移到正式集合
                const completedNewProducts = await Product.find();
                completedNewProducts.sort((a, b) => a.商品編號.localeCompare(b.商品編號));
                await NewProduct.insertMany(completedNewProducts);
                await Product.collection.drop(); // 刪除暫存集合
                console.log(`已將 ${completedNewProducts.length} 筆資料從暫存集合轉移至正式集合: ${newCollectionName}，並已刪除暫存集合。`);
            }

            return res.json(pendingProducts); // 傳回待使用者填寫的產品資訊

        }

    } catch (error) {
        console.error('建立盤點資料庫時發生錯誤:', error);
        if (!res.headersSent) {
            return res.status(500).json({ message: '處理請求時發生錯誤', error: error.message });
        }
    }
});
// API 端點：儲存補齊的新品
app.post('/api/saveCompletedProducts/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
    const collectionTmpName = `${year}${formattedMonth}${storeName}_tmp`; // 暫存集合
    const Product = mongoose.model(collectionName, productSchema);
    const ProductTemp = mongoose.model(collectionTmpName, productSchema); // 暫存集合模型
    const message = '請將頁面重新整理以更新品項';
    const inventoryItems = req.body; // 獲取請求中的數據

    try {
        // 驗證 inventoryItems 是否為陣列且不為空
        if (!Array.isArray(inventoryItems) || inventoryItems.length === 0) {
            return res.status(400).json({ message: '請求體中沒有有效的產品數據' });
        }

        // 使用 Promise.all 並行處理每個項目的更新
        const updateResults = await Promise.all(inventoryItems.map(item => {
            if (!item.商品編號 || item.庫別 === undefined || item.廠商 === undefined) {
                console.warn(`缺少必要欄位，跳過更新商品編號: ${item.商品編號}`);
                return Promise.resolve({ matchedCount: 0, modifiedCount: 0 }); // 跳過並返回成功，避免 Promise.all 失敗
            }
            return ProductTemp.updateOne(
                { 商品編號: item.商品編號 }, // 根據 '商品編號' 更新
                { $set: { 庫別: item.庫別, 廠商: item.廠商, 停用: item.停用} }
            );
        }));

        // 檢查是否有任何更新失敗
        const totalModified = updateResults.reduce((sum, result) => sum + result.modifiedCount, 0);
        console.log(`成功更新 ${totalModified} 個產品`);

        const completedTmpProducts = await ProductTemp.find(); // 取得暫存區數據
        const allProducts = [...completedTmpProducts];

        // 根据商品编号进行排序
        allProducts.sort((a, b) => a.商品編號.localeCompare(b.商品編號));

        if (allProducts.length > 0) {
            // 將所有產品資訊存入資料庫
            await Product.insertMany(allProducts);

            // 刪除整個暫存集合
            await ProductTemp.collection.drop(); // 使用 drop() 方法刪除暫存集合

            // io.to(storeName).emit('newAnnouncement', { message, storeName }); // 考慮是否需要發送通知
            return res.status(201).json({ message: '所有新產品已成功保存並更新正式庫存' });
        } else {
            return res.status(200).json({ message: '沒有需要保存的新產品' }); // 更精確的訊息和狀態碼
        }

    } catch (error) {
        console.error('儲存產品時出錯:', error);
        return res.status(500).json({ message: '儲存失敗', error: error.message });
    }
});
// API端點: 檢查伺服器內部狀況
app.get('/api/checkConnections', (req, res) => {
    // 檢查服務器內部狀況，假設這裡始終有效
    res.status(200).json({ serverConnected: true });
});


const net = require('net');

// API 端點: 檢查EPOS伺服器內部狀況
app.get('/api/ping', (req, res) => {
    const client = new net.Socket();
    client.setTimeout(5000);

    client.connect(443, 'google.com', () => {
        // 連線成功
        res.status(200).json({ eposConnected: true });
        client.destroy();
    });

    client.on('error', (err) => {
        console.error('Connection error:', err);
        res.send({ eposConnected: false });
    });

    client.on('timeout', () => {
        console.error('Connection timeout');
        res.send({ eposConnected: false });
    });
});


app.get(`/api/products`, limiter, async (req, res) => {
    return res.status(100).json({ message: '請選擇門市' }); // 當商店名稱未提供時回覆訊息
});

// 獲取產品數據的 API
app.get(`/api/products/:storeName`, async (req, res) => {
    const storeName = req.params.storeName || 'NA'; // 取得 URL 中的 storeName

    try {
        if (storeName === 'NA') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤，因為請求參數有誤
        } else {

            const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const products = await Product.find(); // 取得產品數據

            // 返回產品數據
            res.json(products);
            res.status(200);

        }
    } catch (error) {
        console.error("取得產品時出錯:", error);
        res.status(500).send('服務器錯誤');
    }

});
// 更新產品數量的 API 端點
app.put('/api/products/:storeName/:productCode/quantity', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店名稱是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 數量 } = req.body;

        // 更新指定產品的數量
        const updatedProduct = await Product.findOneAndUpdate(
            { 商品編號: productCode },
            { 數量 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }

        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productUpdated', updatedProduct);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新產品時出錯:', error);
        res.status(400).send('更新失敗');
    }
});
// 更新產品停用的 API 端點
app.put('/api/products/:storeName/:productCode/depot', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店名稱是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 停用 } = req.body;
        const storeRoom = req.params.storeName;
        // 更新指定產品的數量
        const updatedProduct = await Product.findOneAndUpdate(
            { 商品編號: productCode },
            { 停用 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }
        if (停用 === true) {
            io.to(storeName).emit('productDepotUpdatedV', updatedProduct, storeRoom);
        } else {
            io.to(storeName).emit('productDepotUpdatedX', updatedProduct, storeRoom);
        }
        res.json(updatedProduct);
    } catch (error) {
        console.error('更新產品時出錯:', error);
        res.status(400).send('更新失敗');
    }
});

// 更新產品到期日的 API 端點
app.put('/api/products/:storeName/:productCode/expiryDate', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店名稱是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 到期日 } = req.body;

        // 更新指定產品的到期日
        const updatedProduct = await Product.findOneAndUpdate(
            { 商品編號: productCode },
            { 到期日: 到期日 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }

        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productExpiryDateUpdated', updatedProduct);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新到期日時出錯:', error);
        res.status(400).send('更新失敗');
    }
});

app.put('/api/products/:storeName/batch-update', async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
    const Product = mongoose.model(collectionName, productSchema); // 使用動態集合名稱
    const message = '請將頁面重新整理以更新品項'
    const inventoryItems = req.body; // 獲取請求中的數據

    try {
        // 使用 Promise.all 並行處理每個項目的更新
        const updatePromises = inventoryItems.map(item => {
            return Product.updateOne(
                { 商品編號: item.商品編號 }, // 根據 '商品編號' 更新
                { $set: { 庫別: item.庫別, 廠商: item.廠商, 停用: false } }
            );
        });

        // 等待所有更新完成
        await Promise.all(updatePromises);
        io.to(storeName).emit('newAnnouncement', { message, storeName });

        res.status(200).send({ message: '更新成功' });
    } catch (error) {
        console.error('更新失敗:', error);
        res.status(500).send({ message: '更新失敗', error: error.message });
    }
});

// API 端點處理盤點歸檔請求
app.post('/api/archive/:storeName', limiter, async (req, res) => {
    try {
        const storeName = req.params.storeName;
        const password = req.body.password;
        const adminPassword = process.env.ADMIN_PASSWORD;


        const decryptedPassword = CryptoJS.AES.decrypt(encryptedPassword, process.env.SECRET_KEY).toString(CryptoJS.enc.Utf8);
        if (decryptedPassword !== adminPassword) {
            return res.status(401).json({ message: '密碼不正確' });
        }


        const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
        const Product = mongoose.model(collectionName, productSchema);
        const products = await Product.find(); // 取得產品數據

        // 將資料儲存到檔案中
        const archiveDir = path.join(__dirname, 'archive');
        const filePath = path.resolve(archiveDir, collectionName);
        if (!filePath.startsWith(archiveDir)) {
            return res.status(403).send('無效的檔案路徑');
        }
        fs.writeFileSync(filePath, JSON.stringify(products, null, 2), 'utf-8');

        // 將資料從資料庫中清除
        await Product.deleteMany();

        res.status(200).send('資料歸檔成功');

    } catch (error) {
        console.error('處理歸檔請求時出錯:', error);
        // 避免重複發送回應
        if (!res.headersSent) {
            res.status(500).send('服務器錯誤');
        }
    }
});
// 更新，根據商店名稱清除庫存數據
app.post('/api/clear/:storeName', limiter, async (req, res) => {
    try {
        const storeName = req.params.storeName; // 取得 URL 中的 storeName
        const password = req.body.password;
        const adminPassword = process.env.ADMIN_PASSWORD;
        const decryptedPassword = CryptoJS.AES.decrypt(encryptedPassword, process.env.SECRET_KEY).toString(CryptoJS.enc.Utf8);
        if (decryptedPassword !== adminPassword) {
            return res.status(401).json({ message: '密碼不正確' });
        }

        const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合名稱
        const Product = mongoose.model(collectionName, productSchema);
        logger.log(collectionName);
        const products = await Product.find(); // 取得產品數據

        // 清除庫存
        await Product.deleteMany();

        message.success('庫存資料已成功清除'); // 成功提示

        res.status(200).send('庫存清除成功'); // 返回成功訊息
    } catch (error) {
        // 避免重複發送回應
        if (!res.headersSent) {
            res.status(500).send('服務器錯誤');
        }
    }
});
// 假設您已經在公告發佈的路由中
app.post('/api/announcement', limiter, async (req, res) => {
    try {
        const { message, storeName } = req.body;

        if (!message || !storeName) {
            return res.status(400).json({ message: '公告內容和商店名稱是必要的' });
        }

        // 廣播公告給特定商店房間
        io.to(storeName).emit('newAnnouncement', { message, storeName });

        res.status(200).json({ message: '公告發布成功' });
    } catch (error) {
        console.error('發佈公告時發生錯誤:', error);
        res.status(500).json({ message: '伺服器錯誤' });
    }
});



// 建立 HTTP 端點和 Socket.IO 服務器
const server = http.createServer(app);
const io = new Server(server, {
    cors: {
        origin: '*', // 確保允許來自特定來源的請求
        methods: ['GET', 'POST', 'PUT', 'OPTIONS'],
    },
  
});
let onlineUsers = {}; // 儲存線上使用者和他們的房間

io.on('connection', (socket) => {
    console.log('用戶上線。');

    // 當使用者加入房間時
    socket.on('joinStoreRoom', (storeName) => {
        socket.join(storeName); // 讓使用者加入指定房間

        // 更新線上使用者數量
        onlineUsers[storeName] = (onlineUsers[storeName] || 0) + 1;

        // 廣播目前線上人數
        const currentUserCount = io.sockets.adapter.rooms.get(storeName)?.size || 0; // 目前房間使用者數
        socket.to(storeName).emit('updateUserCount', currentUserCount); // 傳送目前人數
        console.log(`使用者加入商店房間：${storeName}，當前人數：${currentUserCount}。`);
    });

    // 當用戶離開時
    socket.on('leaveStoreRoom', () => {
        console.log('用戶離線。');

        // 尋找使用者目前所在的房間
        const rooms = Object.keys(socket.rooms);
        rooms.forEach((room) => {
            if (onlineUsers[room]) {
                onlineUsers[room] -= 1; // 減少對應房間的線上使用者數
                // 若房間內無線上使用者, 可以選擇刪除房間訊息
                if (onlineUsers[room] <= 0) {
                    delete onlineUsers[room];
                } else {
                    socket.to(room).emit('updateUserCount', onlineUsers[room]); // 廣播更新後的線上人數
                }
            }
            console.log(`使用者離開商店房間：${room}，當前人數：${onlineUsers[room] || 0}。`);
        });
    });

    // 偵測用戶斷開連接
    socket.on('disconnect', () => {
        console.log('用戶已斷開連接。');
        const rooms = Object.keys(socket.rooms);
        rooms.forEach((room) => {
            if (onlineUsers[room]) {
                onlineUsers[room] -= 1; // 減少對應房間的線上使用者數
                // 若房間內無線上使用者, 可以選擇刪除房間訊息
                if (onlineUsers[room] <= 0) {
                    delete onlineUsers[room];
                } else {
                    socket.to(room).emit('updateUserCount', onlineUsers[room]); // 廣播更新後的線上人數
                }
            }
        });
    });
});
// 起動伺服器
const PORT = process.env.PORT || 4000
server.listen(PORT, () => {
    console.log(`伺服器正在連接埠 ${PORT} 上運行`);
});
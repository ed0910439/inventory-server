const express = require('express');
const { Server } = require('socket.io');
const cookieParser = require('cookie-parser');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const multer = require('multer');
const xml2js = require('xml2js');
const net = require('net');
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


const upload = multer({ storage: multer.memoryStorage() }); // 使用內存存儲，方便直接獲取buffer

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
    品號: { type: String, required: true },
    廠商: { type: String, required: false },
    品名: { type: String, required: false },
    規格: { type: String, required: false },
    盤點單位: { type: String, required: false },
    本月報價: { type: String, required: false },
    保存期限: { type: String, required: false },
    本月進貨: { type: String, required: false },
    進貨單位: { type: String, required: false },
    期初盤點: { type: Number, required: false },
    盤點量1: { type: Number, default: undefined },
    盤點量2: { type: Number, default: undefined },
    期末盤點: { type: String, required: false },
    調出: { type: String, required: false },
    調入: { type: String, required: false },
    本月使用量: { type: Number, required: false },
    本月食材成本: { type: Number, required: false },
    本月萬元用量: { type: Number, required: false },
    週用量: { type: Number, required: false },
    盤點日期: { type: String, required: false },

});

// 定義 sanitizeInput 函數
const sanitizeInput = (input) => {
    return encodeURIComponent(input.trim());
};
// 動態產生集合品名
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

const currentlyEditingProductsByStore = {};

app.get('/api/testInventoryTemplate/:storeName', (req, res) => {
    const storeName = req.params.storeName;

    // 返回用于测试的示例数据
    const mockInventoryData = [
        { 停用: 'ture', 品號: '001', 品名: '項目A', 規格: '', 廠商: '待設定', 庫別: '待設定' },
        { 停用: 'false', 品號: '002', 品名: '項目B', 規格: '', 廠商: '待設定', 庫別: '待設定' },
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
            const collectionName = `${year}${formattedMonth}${storeName}_tmp`; // 根據年份、月份和門市產生暫存集合品名
            const newCollectionName = `${year}${formattedMonth}${storeName}`; // 當月正式集合品名
            const lastCollectionName = `${lastYear}${formattedLastMonth}${storeName}`; // 上個月集合品名
            const Product = mongoose.model(collectionName, productSchema); // 暫存模型
            const NewProduct = mongoose.model(newCollectionName, productSchema); // 當月正式模型
            const firstUrl = process.env.FIRST_URL.replace('${today}', today); // 替換 URL 中的變數
            const secondUrl = process.env.SECOND_URL;
            console.log('第一份 HTML URL:', firstUrl);
            console.log('第二份 HTML URL:', secondUrl);

            await Product.deleteMany(); // 清空暫存集合
            console.log(`已清空暫存集合: ${collectionName}`);
            console.log('當月暫存集合品名:', collectionName);
            console.log('上個月集合品名:', lastCollectionName);
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
                        品號: row[9],
                        品名: row[10],
                        規格: row[11],
                    };
                    if (product.模板名稱 === '段純貞' && product.品號) { // 確保有品號
                        newProductsFromHtml.push(product); // 只保存有效的段純貞產品
                    }
                }
            });
            console.log(`從第一個資料來源抓取到 ${newProductsFromHtml.length} 個段純貞產品`);

            // 取得上個月的盤點資料並建立品號索引
            let lastMonthInventory = [];
            const lastMonthInventoryMap = new Map();
            try {
                const sourceCollection = mongoose.connection.collection(lastCollectionName);
                lastMonthInventory = await sourceCollection.find({}).toArray(); // 取得上個月集合所有資料
                lastMonthInventory.forEach(item => lastMonthInventoryMap.set(item.品號, item));
                console.log(`成功取得上個月盤點資料，共 ${lastMonthInventory.length} 筆`);
            } catch (error) {
                console.warn(`無法找到上個月的盤點集合: ${lastCollectionName}，將視為首次盤點。`);
                lastMonthInventory = []; // 如果上個月集合不存在，視為空陣列
            }

            // 建立本月盤點所需的基礎資料
            const refinedData = newProductsFromHtml.map(htmlProduct => {
                const lastMonthData = lastMonthInventoryMap.get(htmlProduct.品號);
                const is停用InHtml = htmlProduct.品名.includes('停用') || htmlProduct.品名.includes('勿下');

                return {
                    停用: lastMonthData?.停用 !== undefined ? lastMonthData.停用 : is停用InHtml,
                    品號: htmlProduct.品號,
                    品名: htmlProduct.品名,
                    規格: lastMonthData?.規格 || '',
                    期末盤點: '',
                    盤點單位: lastMonthData?.盤點單位 || '', // 暫不處理
                    保存期限: '',
                    廠商: lastMonthData?.廠商 || (htmlProduct.品號.includes('KO') || htmlProduct.品號.includes('KL') ? '王座(用)' : htmlProduct.品號.includes('KM') ? '央廚' : '待設定'),
                    庫別: lastMonthData?.庫別 || (lastMonthData?.停用 ? '未使用' : is停用InHtml ? '未使用' : '待設定'),
                    盤點日期: '',
                    期初盤點: lastMonthData?.期末盤點 || '無數據',
                };
            });

            if (refinedData.length > 0) {
                await Product.insertMany(refinedData); // 存入暫存集合
                console.log(`已將 ${refinedData.length} 筆基礎產品資訊存入暫存集合: ${collectionName}`);
            } else {
                console.log('沒有需要新增的產品。');
            }

            // 建立暫存集合的品號映射
            const tempProductCodeMap = new Map(refinedData.map(item => [item.品號, item]));

            // 從第二個 HTML 資料來源抓取資料並更新單位
            console.log(`抓取第二個 HTML 資料...`);
            const secondResponse = await axios.get(secondUrl);
            const secondHtml = secondResponse.data;
            const $second = cheerio.load(secondHtml);

            const secondInventoryData = [];
            $second('table tr').each((i, el) => {
                if (i === 0) return; // 忽略表頭
                const row = $second(el).find('td').map((j, cell) => $second(cell).text().trim()).get();

                if (row.length > 3 && row[0] && row[3]) { // 確保品號和單位存在
                    secondInventoryData.push({
                        品號: row[0],
                        盤點單位: row[3],
                    });
                }
            });
            console.log(`從第二個資料來源抓取到 ${secondInventoryData.length} 筆產品單位資訊`);

            // 更新暫存集合中的產品單位
            secondInventoryData.forEach(item => {
                const tempProduct = tempProductCodeMap.get(item.品號);
                if (tempProduct) {
                    tempProduct.盤點單位 = item.盤點單位;
                }
            });

            const updatedProducts = Array.from(tempProductCodeMap.values());

            // 傳回所有庫別為「待設定」新品項，等待使用者填寫
            const pendingProducts = updatedProducts.filter(product => product.庫別 === '待設定');

            if (pendingProducts.length === 0) {
                // 沒有待設定產品，將暫存資料轉移到正式集合
                const completedNewProducts = await Product.find();
                completedNewProducts.sort((a, b) => a.品號.localeCompare(b.品號));
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
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
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
            if (!item.品號 || item.庫別 === undefined || item.廠商 === undefined) {
                console.warn(`缺少必要欄位，跳過更新品號: ${item.品號}`);
                return Promise.resolve({ matchedCount: 0, modifiedCount: 0 }); // 跳過並返回成功，避免 Promise.all 失敗
            }
            return ProductTemp.updateOne(
                { 品號: item.品號 }, // 根據 '品號' 更新
                { $set: { 庫別: item.庫別, 廠商: item.廠商, 停用: item.停用} }
            );
        }));

        // 檢查是否有任何更新失敗
        const totalModified = updateResults.reduce((sum, result) => sum + result.modifiedCount, 0);
        console.log(`成功更新 ${totalModified} 個產品`);

        const completedTmpProducts = await ProductTemp.find(); // 取得暫存區數據
        const allProducts = [...completedTmpProducts];

        // 根据商品编号进行排序
        allProducts.sort((a, b) => a.品號.localeCompare(b.品號));

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

// 新的 API 端點，處理上傳的 Excel 檔案
app.post('/api/uploadInventory/:storeName', upload.single('inventoryFile'), async (req, res) => {
    console.log('接收的請求:', req.body); // 打印請求體
    console.log('請求文件:', req.file); // 打印上傳的文件
    const storeName = req.params.storeName;
    if (!req.file) {
        return res.status(400).json({ message: '請上傳 Excel 檔案' });
    }

    // 獲取上傳的文件名稱
    const uploadedFileName = req.file.originalname;
    console.log('上傳的文件名:', uploadedFileName);
try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const worksheet = workbook.getWorksheet('總表');
    if (!worksheet) {
        return res.status(400).json({ message: '工作表「總表」不存在' });
    }

    const data = [];
    worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 2) { // 從第三行開始讀取資料
        const rowData = row.values.slice(1); // 去掉第一個空項
        for (let col = 0; col < rowData.length; col++) {
            const cellValue = rowData[col];

            // 檢查cellValue是否是對象，並提取文本
            if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
                // 這裡假設您只需要元件中的text部分
                rowData[col] = cellValue.richText.map(item => item.text).join(''); // 將所有文本合併
            } else {
                rowData[col] = String(cellValue || ''); // 其他情況下，轉為字串
            }
        }
        data.push(rowData); // 存儲處理過的數據
        }
    });
            const collectionName = `${year}${formattedMonth}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);
    const bulkOps = []; // 初始化bulkOps數組
      bulkOps.push({
    updateMany: {
      filter: {}, // 匹配所有文檔
      update: { $set: { 本月報價: 0 } },
    },
  });
    // 更新數據
    const rows = data; // 如果data已包含所有行，則可以直接操作
    rows.forEach(row => {
        const 品號 = String(row[0] || ''); // 假設品號在第1列 (1-based index)
bulkOps.push({
    updateOne: {
        filter: { 品號: String(row[0] || '') }, // 確保品號是字串
        update: {
            $set: {
                廠商: String(row[1]||'未知'),            // 第2欄
                規格: String(row[3]||'未知'),            // 第4欄
                盤點單位: String(row[4]||'未知'),        // 第5欄
                本月報價: parseFloat(row[5]) || 0, // 第6欄
                進貨單位: String(row[8]||'未知'),        // 第9欄
            },
        },
        upsert: false, // 如果品號不存在則不新增
    },
});

    });

    // 如果bulkOps有數據，則進行資料庫更新
    if (bulkOps.length > 0) {
        const result = await Product.bulkWrite(bulkOps);
        res.status(200).json({ message: `成功更新 ${result.modifiedCount} 筆資料` });
    } else {
        res.status(200).json({ message: '沒有找到可更新的品號' });
    }
} catch (error) {
    console.error('處理 Excel 檔案時發生錯誤:', error);
    res.status(500).json({ message: '處理檔案時發生錯誤', error: error.message });
}
});


// 新的 API 端點，處理上傳的本月進貨量 Excel 檔案
app.post('/api/uploadMonthlyPurchase/:storeName', upload.single('monthlyPurchaseFile'), async (req, res) => {
    const storeName = req.params.storeName;
          const collectionName = `${year}${formattedMonth}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);
    if (!req.file) {
        return res.status(400).json({ message: '請上傳本月進貨量文件' });
    }

    try {
        const parser = new xml2js.Parser();
        const result = await parser.parseStringPromise(req.file.buffer.toString('utf-8'));

        const worksheet = result.Workbook.Worksheet[0].Table[0].Row;

        if (!worksheet || worksheet.length < 1) {
            return res.status(400).json({ message: 'XML 檔案格式不正確，缺少資料' });
        }

        const bulkOps = [];
        bulkOps.push({
            updateMany: {
              filter: {}, // 匹配所有文檔
              update: { $set: { 本月進貨: 0 } },
            },
          });
        // 從第二行開始讀取數據，忽略最後的合計行
        for (let i = 1; i < worksheet.length - 1; i++) {
            const rowData = worksheet[i].Cell.map(cell => (cell.Data ? cell.Data[0]._ : '')); // 提取每行的數據
            
            const productCode = rowData[2]; // 第三欄是商品編碼
            const receivedQuantity = rowData[9]; // 第九欄是驗收數量

            if (typeof productCode === 'string') {
                const trimmedProductCode = productCode.trim();

                bulkOps.push({
                    updateOne: {
                        filter: { 品號: trimmedProductCode },
                        update: { $set: { 本月進貨: parseFloat(receivedQuantity) || 0 } },
                        upsert: true,
                    },
                });
            } else {
                console.warn(`無效的商品編碼：${productCode}`);
            }
        }

        // 執行批量更新操作
        if (bulkOps.length > 0) {
            const result = await Product.bulkWrite(bulkOps);
            res.status(200).json({ message: `成功更新 ${result.modifiedCount} 筆本月進貨量` });
        } else {
            res.status(200).json({ message: '沒有找到可更新的商品編碼' });
        }

    } catch (error) {
        console.error('處理本月進貨量 XML 檔案時發生錯誤:', error);
        res.status(500).json({ message: '處理檔案時發生錯誤', error: error.message });
    }
});

// API端點: 檢查伺服器內部狀況
app.get('/api/checkConnections', (req, res) => {
    // 檢查服務器內部狀況，假設這裡始終有效
    res.status(200).json({ serverConnected: true });
});



// API 端點: 檢查EPOS伺服器內部狀況
app.get('/api/ping', (req, res) => {
    const client = new net.Socket();
    client.setTimeout(5000);

    client.connect(8090, 'epos.kingza.com.tw', () => {
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
    return res.status(100).json({ message: '請選擇門市' }); // 當商店品名未提供時回覆訊息
});

// 獲取產品數據的 API
app.get(`/api/products/:storeName`, async (req, res) => {
    const storeName = req.params.storeName || 'NA'; // 取得 URL 中的 storeName

    try {
        if (storeName === 'NA') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤，因為請求參數有誤
        } else {

            const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
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
// 更新產品數量1的 API 端點
app.put('/api/products/:storeName/:productCode/quantity1', limiter, async (req, res) => {
    const { storeName, productCode } = req.params;
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);
    const storeRoom = req.params.storeName;

    const { 盤點量1 } = req.body; // 從請求體中獲取期末盤點

    if (typeof 盤點量1 === 'undefined' || 盤點量1 === null) {
        return res.status(400).json({ message: '期末盤點數量是必需的' });
    }

    try {
        let updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 盤點量1: 盤點量1 } }, // 使用 $set 操作符確保只更新這個字段
            { new: true } // 返回更新後的文檔
        );

        if (!updatedProduct) {
            return res.status(404).json({ message: '產品未找到或門市名稱不匹配' });
        }

        // 計算新的期末盤點 (合計值)
        const new期末盤點 = (updatedProduct.盤點量1 || 0) + (updatedProduct.盤點量2 || 0);

        // 將新的期末盤點更新回資料庫
        updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 期末盤點: new期末盤點 } },
            { new: true } // 再次返回更新後的文檔，確保包含最新的期末盤點
        );

        if (!updatedProduct) {
            // 這應該不太可能發生，因為上一步已經找到並更新了
            return res.status(404).json({ message: '更新期末盤點失敗，產品未找到' });
        }

        // 廣播更新訊息給所有用戶，現在 updatedProduct 包含了最新的 期末盤點 (合計值)
        if (typeof io !== 'undefined') { // 確保 io 存在
            io.to(storeName).emit('productUpdated', updatedProduct, storeRoom);
        }
        if (currentlyEditingProductsByStore[storeName] && currentlyEditingProductsByStore[storeName][productCode]) {
            delete currentlyEditingProductsByStore[storeName][productCode];
            // 廣播此產品的編輯狀態已停止
            io.to(storeName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
            console.log(`產品 ${productCode} 在門市 ${storeName} 的編輯狀態已在後端清除。`);
        }

        res.status(200).json(updatedProduct);
    } catch (error) {
        console.error(`更新產品 ${productCode} 的盤點量1時發生錯誤:`, error);
        res.status(500).json({ message: '伺服器錯誤，無法更新盤點量1的數量' });
    }
});
// 更新產品數量2的 API 端點
app.put('/api/products/:storeName/:productCode/quantity2', limiter, async (req, res) => {
    const { storeName, productCode } = req.params;
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);
    const storeRoom = req.params.storeName;

    const { 盤點量2 } = req.body; // 從請求體中獲取期末盤點

    if (typeof 盤點量2 === 'undefined' || 盤點量2 === null) {
        return res.status(400).json({ message: '期末盤點數量是必需的' });
    }

    try {
        let updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 盤點量2: 盤點量2 } }, // 使用 $set 操作符確保只更新這個字段
            { new: true } // 返回更新後的文檔
        );

        if (!updatedProduct) {
            return res.status(404).json({ message: '產品未找到或門市名稱不匹配' });
        }

        // 計算新的期末盤點 (合計值)
        const new期末盤點 = (updatedProduct.盤點量1 || 0) + (updatedProduct.盤點量2 || 0);

        // 將新的期末盤點更新回資料庫
        updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 期末盤點: new期末盤點 } },
            { new: true } // 再次返回更新後的文檔，確保包含最新的期末盤點
        );

        if (!updatedProduct) {
            // 這應該不太可能發生，因為上一步已經找到並更新了
            return res.status(404).json({ message: '更新期末盤點失敗，產品未找到' });
        }

        // 廣播更新訊息給所有用戶，現在 updatedProduct 包含了最新的 期末盤點 (合計值)
        if (typeof io !== 'undefined') { // 確保 io 存在
            io.to(storeName).emit('productUpdated', updatedProduct, storeRoom);
        }
        if (currentlyEditingProductsByStore[storeName] && currentlyEditingProductsByStore[storeName][productCode]) {
            delete currentlyEditingProductsByStore[storeName][productCode];
            // 廣播此產品的編輯狀態已停止
            io.to(storeName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
            console.log(`產品 ${productCode} 在門市 ${storeName} 的編輯狀態已在後端清除。`);
        }

        res.status(200).json(updatedProduct);
    } catch (error) {
        console.error(`更新產品 ${productCode} 的盤點量2時發生錯誤:`, error);
        res.status(500).json({ message: '伺服器錯誤，無法更新盤點量2的數量' });
    }
});
// 更新產品停用的 API 端點
app.put('/api/products/:storeName/:productCode/depot', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 停用 } = req.body;
        const storeRoom = req.params.storeName;
        // 更新指定產品的數量
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
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

// 更新產品保存期限的 API 端點
app.put('/api/products/:storeName/:productCode/expiryDate', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 保存期限 } = req.body;

        // 更新指定產品的保存期限
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { 保存期限: 保存期限 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }

        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productExpiryDateUpdated', updatedProduct);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新保存期限時出錯:', error);
        res.status(400).send('更新失敗');
    }
});

// 更新產品廠商 API 端點
app.put('/api/products/:storeName/:productCode/vendor', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 廠商 } = req.body;

        // 更新指定產品的廠商
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { 廠商 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }

        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productVendorUpdated', updatedProduct);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新保存期限時出錯:', error);
        res.status(400).send('更新失敗');
    }
});
// 更新產品庫別 API 端點
app.put('/api/products/:storeName/:productCode/layer', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 庫別 } = req.body;

        // 更新指定產品的庫別
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { 庫別 },
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }

        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productLayerUpdated', updatedProduct);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新保存期限時出錯:', error);
        res.status(400).send('更新失敗');
    }
});

app.put('/api/products/:storeName/batch-update', async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema); // 使用動態集合品名
    const message = '請將頁面重新整理以更新品項'
    const inventoryItems = req.body; // 獲取請求中的數據

    try {
        // 使用 Promise.all 並行處理每個項目的更新
        const updatePromises = inventoryItems.map(item => {
            return Product.updateOne(
                { 品號: item.品號 }, // 根據 '品號' 更新
                { $set: { 庫別: item.庫別, 廠商: item.廠商, 停用: item.停用 } }
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


        const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
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
// 更新，根據商店品名清除庫存數據
app.post('/api/clear/:storeName', limiter, async (req, res) => {
    try {
        const storeName = req.params.storeName; // 取得 URL 中的 storeName
        const password = req.body.password;
        const adminPassword = process.env.ADMIN_PASSWORD;
        const decryptedPassword = CryptoJS.AES.decrypt(encryptedPassword, process.env.SECRET_KEY).toString(CryptoJS.enc.Utf8);
        if (decryptedPassword !== adminPassword) {
            return res.status(401).json({ message: '密碼不正確' });
        }

        const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
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
            return res.status(400).json({ message: '公告內容和商店品名是必要的' });
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
const onlineUsers = {}; // 保持您現有的在線用戶計數
// Socket.io 連接處理
io.on('connection', (socket) => {
    console.log(`用戶連接: ${socket.id}`);

    // 用戶加入商店房間
    socket.on('joinStoreRoom', (storeName) => {
        let currentStoreName = socket.data.storeName;
        if (currentStoreName && currentStoreName !== storeName) {
            // 如果用戶從其他房間切換過來，先離開舊房間
            socket.leave(currentStoreName);
            console.log(`使用者 ${socket.id} 離開商店房間：${currentStoreName}`);
            // 清除該用戶在舊房間的編輯狀態
            if (currentlyEditingProductsByStore[currentStoreName]) {
                for (const productCode in currentlyEditingProductsByStore[currentStoreName]) {
                    if (currentlyEditingProductsByStore[currentStoreName][productCode].by === socket.id) {
                        delete currentlyEditingProductsByStore[currentStoreName][productCode];
                        // 廣播此產品的編輯狀態已停止
                        io.to(currentStoreName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
                        console.log(`用戶 ${socket.id} 在離開房間時，清除產品 ${productCode} 的編輯狀態。`);
                    }
                }
            }
        }

        socket.join(storeName);
        socket.data.storeName = storeName; // 儲存當前房間名稱
        console.log(`使用者 ${socket.id} 加入商店房間：${storeName}，當前人數：${io.sockets.adapter.rooms.get(storeName)?.size || 0}。`);

        // 發送當前所有正在編輯的產品狀態給新加入的客戶端
        if (currentlyEditingProductsByStore[storeName]) {
            socket.emit('currentEditingState', currentlyEditingProductsByStore[storeName]);
        }
        // 更新房間人數
        io.to(storeName).emit('updateUserCount', io.sockets.adapter.rooms.get(storeName)?.size || 0);
    });

    // 監聽產品開始編輯事件
    socket.on('startEditingProduct', ({ productCode, storeName, quantityType }) => { // <--- 這裡必須有 quantityType
        if (!currentlyEditingProductsByStore[storeName]) {
            currentlyEditingProductsByStore[storeName] = {};
    }
    // 儲存編輯狀態，包含編輯類型
    currentlyEditingProductsByStore[storeName][productCode] = {
        by: socket.id,
        timestamp: Date.now(),
        editingType: quantityType // 保存編輯的類型 (quantity1 或 quantity2)
    };
    // 廣播編輯狀態更新，包含編輯類型
    io.to(storeName).emit('productEditingStateUpdate', {
        productCode,
        status: 'editing',
        by: socket.id,
        editingType: quantityType // 傳遞編輯類型給前端
    });
    console.log(`廣播產品 ${productCode} 的 ${quantityType} 編輯狀態開始，由 ${socket.id} 編輯`);
});

    socket.on('stopEditingProduct', ({ productCode, storeName, quantityType }) => { // 接收 quantityType
        if (currentlyEditingProductsByStore[storeName] &&
            currentlyEditingProductsByStore[storeName][productCode] &&
            currentlyEditingProductsByStore[storeName][productCode].by === socket.id &&
            currentlyEditingProductsByStore[storeName][productCode].editingType === quantityType // 也要匹配類型
        ) {
            delete currentlyEditingProductsByStore[storeName][productCode];
            // 廣播停止編輯狀態，這裡 status 為 'idle'，不需要傳遞 editingType，因為整個產品的編輯狀態都清除了
            io.to(storeName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
            console.log(`廣播產品 ${productCode} 的 ${quantityType} 編輯狀態停止，由 ${socket.id} 停止`);
        }
    });


    // 處理斷開連接
    socket.on('disconnect', () => {
        console.log(`用戶斷開: ${socket.id}`);

        let currentStoreName = socket.data.storeName;
            if (currentStoreName && currentlyEditingProductsByStore[currentStoreName]) {
        for (const productCode in currentlyEditingProductsByStore[currentStoreName]) {
            // 注意：這裡斷開連接時，我們不能判斷是哪個 quantityType，
            // 只能清除該用戶編輯的所有產品編輯狀態。
            // 由於一個產品可能只有一個編輯鎖定，所以直接清除即可
            if (currentlyEditingProductsByStore[currentStoreName][productCode].by === socket.id) {
                console.log(`用戶 ${socket.id} 斷開連接時，清除產品 ${productCode} 的編輯狀態。`);
                delete currentlyEditingProductsByStore[currentStoreName][productCode];
                io.to(currentStoreName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
            }
        }

        
            // 更新離線用戶的在線人數（如果需要）
            onlineUsers[currentStoreName] = (onlineUsers[currentStoreName] || 1) - 1;
            io.to(currentStoreName).emit('updateUserCount', io.sockets.adapter.rooms.get(currentStoreName)?.size || 0);
            console.log(`使用者 ${socket.id} 離開商店房間：${currentStoreName}，當前人數：${io.sockets.adapter.rooms.get(currentStoreName)?.size || 0}。`);
        }
    });

    // 定時器檢查並清除過期的編輯狀態
    // 這個定時器需要在每個 storeName 下獨立檢查
    setInterval(() => {
        const fiveMinutesAgo = Date.now() - (5 * 60 * 1000); // 5 分鐘前
        for (const storeName in currentlyEditingProductsByStore) {
            const storeEditingProducts = currentlyEditingProductsByStore[storeName];
            for (const productCode in storeEditingProducts) {
                // 無需檢查 editingType，只要過期就清除
                if (storeEditingProducts[productCode].timestamp < fiveMinutesAgo) {
                    console.log(`自動清除產品 ${productCode} 在門市 ${storeName} 的過期編輯狀態`);
                    delete storeEditingProducts[productCode];
                    io.to(storeName).emit('productEditingStateUpdate', { productCode, status: 'idle' });
                }
            }
        }
    }, 30 * 1000); // 每 30 秒檢查一次
});

// 起動伺服器
const PORT = process.env.PORT || 4000
server.listen(PORT, () => {
    console.log(`伺服器正在連接埠 ${PORT} 上運行`);
});
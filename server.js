// server.js
const express = require('express');
const { Server } = require('socket.io'); // 引入 Socket.IO Server
const cookieParser = require('cookie-parser');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const http = require('http'); // 引入 http 模組
const ExcelJS = require('exceljs');
const axios = require('axios');
const cheerio = require('cheerio');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const csrf = require('csurf');
const bodyParser = require('body-parser');

require('dotenv').config(); // 環境變數管理

const app = express(); // 初始化 Express 應用

// 中介配置
app.use(cookieParser());
app.use(cors());
app.use(bodyParser.urlencoded({ extended: true })); // 解析 URL 編碼的請求
app.use(bodyParser.json()); // 解析 JSON 請求
app.use(helmet()); // 使用 Helmet 增加安全性

// 設定 CSRF 保護
const csrfProtection = csrf({ cookie: true }); // 使用 cookie 存儲 CSRF 令牌
app.use(csrfProtection); // 使用 CSRF 中介

// 提供 CSRF 令牌的 API 端點
app.get('/api/csrf-token', (req, res) => {
    res.json({ csrfToken: req.csrfToken() }); // 回傳 CSRF 令牌
});

// CSRF 錯誤處理
app.use((err, req, res, next) => {
    if (err.code === 'EBADCSRFTOKEN') {
        return res.status(403).send('CSRF token validation failed'); // 返回 CSRF 錯誤
    }
    next(err); // 繼續處理其他錯誤
});

const archiveLimiter = rateLimit({
    windowMs: 1 * 60 * 1000, // 1 分鐘
    max: 5, // 每個 IP 每窗口限制 5 次請求
});

// 連接到 MongoDB
mongoose.connect(process.env.MONGODB_URI, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
})
.then(() => console.log('成功連接到 MongoDB'))
.catch(err => console.error('MongoDB 連接錯誤:', err));

// 定義產品模型
const productSchema = new mongoose.Schema({
    商品編號: { type: String, required: true },
    商品名稱: { type: String, required: false },
    規格: { type: String, required: false },
    數量: { type: Number, required: false },
    單位: { type: String, required: false },
    到期日: { type: String, required: false },
    廠商: { type: String, required: false },
    庫別: { type: String, required: false },
    盤點日期: { type: String, required: false },
    期初庫存: { type: String, required: false },
});

// 動態生成集合名稱
const currentDate = new Date();
let year = currentDate.getFullYear();
let latesrmonth = String(currentDate.getMonth()).padStart(2, '0');
let month = String(currentDate.getMonth() + 1).padStart(2, '0'); // 注意：月份從0開始，因此需要加1
let day = String(currentDate.getDate()).padStart(2, '0');

// 根據日期決定使用的月份
if (day < 16) {
    month -= 1; // 回到上個月
    if (month === 0) {
        month = 12; // 回到前一年的12月
        year -= 1;
    }
}

// 盤點開始的 API 端點
app.get('/api/startInventory/:storeName', archiveLimiter, async (req, res) => {
    const storeName = sanitizeInput(req.params.storeName);
    console.log(`獲取庫存的門市名稱: ${storeName}`);

    if (!storeName || storeName === 'notStart') {
        console.error('門市錯誤: 未提供有效的店名');
        return res.status(400).json({ message: '門市錯誤' });         // 使用 400 Bad Request 返回錯誤
    }

    try {
        const today = `${year}-${month}-${day}`;
        const collectionName = `${year}${month}${storeName}`;
        const latesCollectionName = `${year}${latesrmonth}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);
        const firstUrl = process.env.FIRST_URL.replace('${today}', today); // 替換 URL 中的變數
        const secondUrl = process.env.SECOND_URL;

        // 抓取第一份 HTML 新資料
        console.log(`抓取 HTML 資料...`);
        const firstResponse = await axios.get(firstUrl);
        const firstHtml = firstResponse.data;
        const $first = cheerio.load(firstHtml);

        const newProducts = [];
        $first('table tr').each((i, el) => {
            if (i === 0) return; // 忽略表头
            const row = $first(el).find('td').map((j, cell) => $first(cell).text().trim()).get();
            if (row.length > 3 && row[1] === '段純貞') {
                const product = {
                    商品編號: row[9],
                    商品名稱: row[10],
                    規格: row[11],
                };
                newProducts.push(product);
            }
        });

        console.log(`從第一個資料源抓取到 ${newProducts.length} 個新產品`);

        // 從第二個 HTML 數據源抓取數據
        console.log(`抓取第二個 HTML 資料...`);
        const secondResponse = await axios.get(secondUrl);
        const secondHtml = secondResponse.data;
        const $second = cheerio.load(secondHtml);
        const secondInventoryData = [];
        
        $second('table tr').each((i, el) => {
            if (i === 0) return; // 忽略表头
            const row = $second(el).find('td').map((j, cell) => $second(cell).text().trim()).get();
            if (row.length > 3) {
                const product = {
                    商品編號: row[0],
                    單位: row[3] || '',
                };
                secondInventoryData.push(product); 
            }
        });

        console.log(`從第二個資料源抓取到 ${secondInventoryData.length} 個舊產品`);

        // 以商品編號映射第二份數據的單位
        const secondInventoryMap = secondInventoryData.reduce((map, item) => {
            map[item.商品編號] = item.單位;
            return map;
        }, {});

        // 更新第一份資料與第二份資料
        newProducts.forEach(product => {
            const matchedUnit = secondInventoryMap[product.商品編號];
            if (matchedUnit) {
                product.單位 = matchedUnit;
            }
        });

        // 獲取源集合數據，進行產品和數據庫的比對
        const sourceCollection = mongoose.connection.collection(latesCollectionName);
        const inventoryData = await sourceCollection.find({}).toArray();
        console.log(`獲取源數據完成, 條目數: ${inventoryData.length}`);

        // 處理上期的盤點數據
        const refinedData = inventoryData.map(item => ({
            商品編號: item.商品編號,
            商品名稱: item.商品名稱,
            規格: item.規格 || '',
            數量: '',
            單位: item.單位 || '',
            到期日: '',
            廠商: item.廠商 || '',
            庫別: item.庫別 || '',
            盤點日期: '',
            期初庫存: item.數量 || ''
        }));

        if (refinedData.length > 0) {
            await Product.insertMany(refinedData);
            console.log(`插入 ${refinedData.length} 條產品數據到集合: ${collectionName}`);
        }

        // 创建映射以便通过商品编号查找
        const inventoryMap = inventoryData.reduce((map, item) => {
            map[item.商品編號] = {
                庫別: item.庫別 || '待設定',
                廠商: item.廠商 || '',
                期初庫存: item.數量 || '無紀錄'
            };
            return map;
        }, {});

        // 更新新產品數據
        const updatedProducts = newProducts.map(product => {
            const sourceData = inventoryMap[product.商品編號];
                        if (sourceData) {
                product.庫別 = sourceData.庫別;
                product.廠商 = sourceData.廠商;
                product.期初庫存 = sourceData.期初庫存;
            } else {
                product.庫別 = '待設定';
            }
            return product;
        });

        // 返回需新設置的商品
        const pendingProducts = updatedProducts.filter(product => product.庫別 === '待設定');

        if (pendingProducts.length > 0) {
            console.log(`找到 ${pendingProducts.length} 個需設置的新品`);
            return res.json(pendingProducts); // 返回待用戶填寫的產品信息
        } else {
            console.log('没有待设置的产品项');
            return res.status(200).json({ message: '没有待设置的产品项' });
        }
    } catch (error) {
        console.error('处理开始盘点请求时出错:', error);
        if (!res.headersSent) {
            return res.status(500).json({ message: '处理请求时出错', error: error.message });
        }
    }
});

// API 端點：保存補齊的新品
app.post('/api/saveCompletedProducts/:storeName', archiveLimiter, async (req, res) => {
    const storeName = sanitizeInput(req.params.storeName) || 'notStart'; // 获取 URL 中的 storeName
    
    try {
        if (storeName === 'notStart') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因为请求参数有误
        } else {
            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const completedProducts = req.body;

            // 驗證每個產品是否包含必填字段
            const validProducts = completedProducts.map(product => ({
                商品編號: product.商品編號,
                商品名稱: product.商品名稱,
                規格: product.規格,
                單位: product.單位,
                廠商: product.廠商 || '未使用', // 如果未選擇，設為'未使用'
                庫別: product.庫別 || '未使用',   // 如果未選擇，設為'未使用'
            }));

            if (validProducts.length > 0) {
                // 將完成的產品信息存入資料庫
                await Product.insertMany(validProducts);
                return res.status(201).json({ message: '所有新產品已成功保存' });
            } else {
                return res.status(400).json({ message: '缺少必填字段，無法保存產品' });
            }
        }
    } catch (error) {
        console.error('保存產品時出錯:', error);
        return res.status(500).json({ message: '保存失敗' });
    }
});

// API 端點處理盤點歸檔請求
app.post('/api/archive/:storeName', archiveLimiter, async (req, res) => {
    try {
        const storeName = sanitizeInput(req.params.storeName);
        const encryptedPassword = req.body.password; 
        const adminPassword = process.env.ADMIN_PASSWORD;

        const decryptedPassword = CryptoJS.AES.decrypt(encryptedPassword, process.env.SECRET_KEY).toString(CryptoJS.enc.Utf8);
        if (decryptedPassword !== adminPassword) {
            return res.status(401).json({ message: '密碼不正確' });
        }

        const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
        const Product = mongoose.model(collectionName, productSchema);
        const products = await Product.find(); // 獲取產品數據

        // 將數據保存到文件中
        const archiveDir = path.join(__dirname, 'archive');
        const filePath = path.resolve(archiveDir, collectionName);
        if (!filePath.startsWith(archiveDir)) {
            return res.status(403).send('無效的文件路徑');
        }
        fs.writeFileSync(filePath, JSON.stringify(products, null, 2), 'utf-8');

        // 將數據從資料庫中清除
        await Product.deleteMany();

        res.status(200).send('數據歸檔成功');

    } catch (error) {
        console.error('處理歸檔請求時出錯:', error);
// 非法請求處理
        if (!res.headersSent) {
            res.status(500).send('伺服器錯誤');
        }
    }
});

// 創建 HTTP 伺服器並加入 Socket.IO
const server = http.createServer(app);
const io = new Server(server, {
    cors: {
        origin: '*', // 允許所有來源的請求
        methods: ['GET', 'POST'],
    },
});

// Socket.IO 連接管理
io.on('connection', (socket) => {
    console.log('使用者上線。');

    // 當用戶加入指定房間
    socket.on('joinStoreRoom', (storeName) => {
        socket.join(storeName); // 將用戶加入房間
        console.log(`使用者加入商店房間：${storeName}`);
        
        // 廣播當前在線人數
        const onlineUsers = io.sockets.adapter.rooms.get(storeName)?.size || 0; // 獲取該房間的用戶數量
        socket.to(storeName).emit('updateUserCount', onlineUsers); // 向房間內的其他用戶發送當前人數
    });

    socket.on('disconnect', () => {
        console.log('使用者離線。');
    });
});

// 啟動伺服器
const PORT = process.env.PORT || 4000;
server.listen(PORT, () => {
    console.log(`伺服器正在端口 ${PORT} 上運行`);
});

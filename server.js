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
//const csrf = require('csurf');
const bodyParser = require('body-parser');
const morgan = require('morgan'); // 新增日誌中介

require('dotenv').config();

const app = express(); 

// 中介配置
app.use(morgan('combined')); // 使用 morgan 記錄 HTTP 請求
app.use(cookieParser());
app.use(cors());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(helmet());
app.enable('trust proxy');

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
app.use(cors({ origin: 'https://inventory-client.edc-pws.com' })); // 或者使用 '*' 来允许所有来源

// 配置 API 請求的速率限制，防止濫用
const archiveLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 分鐘窗口
    max: 100, // 每個 IP 15 分鐘內最多可以請求 100 次
});
app.use('/api/', archiveLimiter); // 只對 API 請求應用 rate limit



// 連接到 MongoDB
mongoose.connect(process.env.MONGODB_URI, {
  ssl: true,
})
.then(() => console.log('成功連接到 MongoDB'))
.catch(err => console.error('MongoDB 連接錯誤:', err));

// 定義產品模型
// 初始化 Express 應用後
const productSchema = new mongoose.Schema({
    商品編號: { type: String, required: true },
    商品名稱: { type: String, required: false },
    規格: { type: String, required: false },
    數量: { type: Number , rquired: false },
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
// 动态生成集合名称
const currentDate = new Date();
let year = currentDate.getFullYear();
let lastYear = currentDate.getFullYear();
let month, lastMonth;

// 獲取當前日期的日
let day = currentDate.getDate();

// 根據日期決定月份
if (day <= 15) {
    // 每月15日（含）以前
    month = currentDate.getMonth(); // 上個月
    lastMonth = currentDate.getMonth() - 1; // 上上一個月

    // 如果上一個月為 -1(即1月)，則需要調整年份
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
const formattedMonth = String(month).padStart(2, '0');
const formattedLastMonth = String(lastMonth).padStart(2, '0'); // 轉換成1-12格式

console.log(year, formattedMonth, day); // 輸出當前年份、月份、日期
console.log(year, formattedLastMonth, day); // 輸出上個月份的年份、月份、日期



app.get('/api/startInventory/:storeName', archiveLimiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 獲取 URL 中的 storeName

    try {
        if (storeName === 'notStart'){
            res.status(204).send('尚未選擇門市'); // 使用 204，
        } else {

            const today = `${year}-${month}-${day}`;
            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const lastCollectionName = `${lastYear}${lastMonth}${storeName}`; // 动态生成集合名称
            const Product = mongoose.model(collectionName, productSchema);
            const firstUrl = process.env.FIRST_URL.replace('${today}', today); // 替換 URL 中的變數
            const secondUrl = process.env.SECOND_URL;
            // 抓取第一份 HTML 新資料
            console.log(collectionName);
            console.log(lastCollectionName);
            console.log(`抓取 HTML 資料...`);

            const firstResponse = await axios.get(firstUrl);
            const firstHtml = firstResponse.data;
            const $first = cheerio.load(firstHtml);

            const newProducts = [];
            $first('table tr').each((i, el) => {
                if (i === 0) return; // 忽略表头
                const row = $first(el).find('td').map((j, cell) => $first(cell).text().trim()).get();

                if (row.length > 3) {
                    const product = {
                        模板名稱: row[1],
                        商品編號: row[9],
                        商品名稱: row[10],
                        規格: row[11],
                    };
                    if (product.模板名稱 == '段純貞') {
                        newProducts.push(product); // 只保存有效的產品
                    }
                }
            });
            console.log(`從第一個資料源抓取到 ${newProducts.length} 個新產品`);

            // 獲取源集合數據進行比對
            const sourceCollection = mongoose.connection.collection(lastCollectionName);
            const inventoryData = await sourceCollection.find({}).toArray(); // 獲取源集合數據

            // 處理最新的盤點數據
            const refinedData = inventoryData.map(item => ({
                商品編號: item.商品編號,
                商品名稱: item.商品名稱,
                規格: item.規格 || '',
                數量: '', // 將數量設置為空
                單位: item.單位 || '',
                到期日: '', // 將到期日設置為空
                廠商: item.廠商 || '',
                庫別: item.庫別 || '',
                盤點日期: '', // 將盤點日期設置為空
                期初庫存: item.數量 || '' // 將數量拷貝到期初庫存
            }));
            if (refinedData.length > 0) {
                // 將完成的產品信息存入資料庫
                await Product.insertMany(refinedData);
            }

            // 创建一个映射，方便通过商品编號查找
            const inventoryMap = {};
            inventoryData.forEach(item => {
                inventoryMap[item.商品編號] = {
                    庫別: item.庫別 || '待設定', // 如果没有则标记為待設置
                    廠商: item.廠商 || '',
                    期初庫存: item.數量 || '無紀錄' // 將數量重命名為期初庫存
                };
            });

            // 更新新產品數據
            const updatedProducts = newProducts.map(product => {
                const sourceData = inventoryMap[product.商品編號]; // 通过商品编號獲取對应數據
                if (sourceData) {
                    // 填入庫别
                    product.庫別 = sourceData.庫別;
                    product.廠商 = sourceData.廠商;
                    product.期初庫存 = sourceData.期初庫存; // 將數量字段重命名為期初庫存
                } else {
                    // 如果没有找到匹配的商品编號，設置庫别為待設置
                    product.庫別 = '待設定';
                }
                return product; // 返回更新後的產品對象
            });

            // 从第二个 HTML 數據源抓取數據
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
                        商品編號: row[0] || '未知',
                        單位: row[3] || '未設定',
                    };
                    if (product.商品編號 && product.單位) {
                        secondInventoryData.push(product); // 將有效的產品添加到列表中
                    }
                }
            });

            // 创建一个映射以比對第二个數據源
            const secondInventoryMap = {};
            secondInventoryData.forEach(item => {
                secondInventoryMap[item.商品編號] = item.單位; // 將单位与商品编號映射
            });

            // 更新產品數據，结合第二个數據源中的单位
            updatedProducts.forEach(product => {
                if (secondInventoryMap[product.商品編號]) {
                    product.單位 = secondInventoryMap[product.商品編號]; // 根据商品编號更新单位
                }
            });

            // 返回所有庫别為“待設定”的新品项，等待用户填写
            const pendingProducts = updatedProducts.filter(product => product.庫別 === '待設定');

            if (pendingProducts.length > 0) {
                return res.json(pendingProducts); // 返回待用户填写的產品信息
            } else {
                console.log('没有待設置的產品项');
                return res.status(200).json({ message: '没有待設置的產品项' });
            }
        }

    } catch (error) {
        console.error('處理開始盤點请求时出错:', error);
        if (!res.headersSent) {
            return res.status(500).json({ message: '處理请求时出错', error: error.message });
        }
    }
});
// API 端點：保存補齊的新品
app.post('/api/saveCompletedProducts/:storeName', archiveLimiter, async (req, res) => {

    const storeName = req.params.storeName|| 'notStart'; // 獲取 URL 中的 storeName

    try {
        if (storeName === 'notStart') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因為请求参數有误
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

// API端點: 檢查伺服器內部狀況
app.get('/api/checkConnections', (req, res) => {
    // 檢查伺服器內部狀況，假設這裡始終有效
    res.status(200).json({ serverConnected: true });
});


const net = require('net');

// API 端點: 檢查EPOS伺服器內部狀況
app.get('/api/ping', (req, res) => {
    const client = new net.Socket();
    client.setTimeout(5000);

    client.connect(443, 'google.com', () => {
        // 连接成功
        res.status(200).json({ eposConnected: true });
        client.destroy();
    });

    client.on('error', (err) => {
        console.error('Connection error:', err);
        res.send({ connected: false });
    });

    client.on('timeout', () => {
        console.error('Connection timeout');
        res.send({ connected: false });
    });
});


app.get(`/api/products`, archiveLimiter, async (req, res) => {
    return res.status(100).json({ message: '請選擇門市' }); // 當商店名稱未提供時回覆消息
    });

// 獲取產品數據的 API
app.get(`/api/products/:storeName`, archiveLimiter, async (req, res) => {
    const storeName = req.params.storeName|| 'notStart'; // 獲取 URL 中的 storeName

    try {
        if (storeName === '') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因為请求参數有误
        } else {

            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const products = await Product.find(); // 獲取產品數據

            // 返回產品數據
            res.json(products);
            res.status(200); 

        }
    } catch (error) {
            console.error("獲取產品時出錯:", error);
            res.status(500).send('伺服器錯誤');
        }
    
});
// 更新產品數量的 API 端點
app.put('/api/products/:storeName/:productCode/quantity', archiveLimiter, async (req, res) => {
        const storeName = req.params.storeName|| 'notStart'; // 獲取 URL 中的 storeName

    try {
        if (storeName === 'notStart') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因為请求参數有误
        } else {

            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const products = await Product.find(); // 獲取產品數據
      const { productCode } = req.params;
      const { 數量 } = req.body;

      // 更新指定產品的數量
      const updatedProduct = await Product.findOneAndUpdate(
          { 商品編號: productCode },
          { 數量: 數量 },
          { new: true }
      );

      if (!updatedProduct) {
          return res.status(404).send('產品未找到');
      }
      // 廣播更新消息给所有用戶
    io.to(storeName).emit('productUpdated', updatedProduct);

      res.json(updatedProduct);
  }} catch (error) {
      console.error('更新產品時出錯:', error);
      res.status(400).send('更新失敗');
  }
});

// 更新產品到期日的 API 端點
app.put('/api/products/:storeName/:productCode/expiryDate', archiveLimiter, async (req, res) => {
        const storeName = req.params.storeName|| 'notStart'; // 獲取 URL 中的 storeName

    try {
        if (storeName === 'notStart') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因為请求参數有误
        } else {

            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const products = await Product.find(); // 獲取產品數據
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
      // 廣播更新消息给所有用戶
      io.to(storeName).emit('productUpdated', updatedProduct);
      
      res.json(updatedProduct);
  }} catch (error) {
      console.error('更新到期日時出錯:', error);
      res.status(400).send('更新失敗');
  }
});

// API 端點處理盤點歸檔請求
app.post('/api/archive/:storeName', archiveLimiter, async (req, res) => {
    try {
        const storeName = req.params.storeName;
        const password = req.body.password;
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
        // 避免重复发送响应
        if (!res.headersSent) {
            res.status(500).send('伺服器錯誤');
        }
    }
});
// 更新，根据商店名称清除库存数据
app.post('/api/clear/:storeName', archiveLimiter, async (req, res) => {
    try {
        const storeName = req.params.storeName; // 获取 URL 中的 storeName
        const password = req.body.password;
        const adminPassword = process.env.ADMIN_PASSWORD;
        const decryptedPassword = CryptoJS.AES.decrypt(encryptedPassword, process.env.SECRET_KEY).toString(CryptoJS.enc.Utf8);
        if (decryptedPassword !== adminPassword) {
            return res.status(401).json({ message: '密碼不正確' });
        }

        const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
        const Product = mongoose.model(collectionName, productSchema);
        logger.log(collectionName);
        const products = await Product.find(); // 獲取產品數據

        // 清除库存
        await Product.deleteMany();

        message.success('库存资料已成功清除'); // 成功提示

        res.status(200).send('库存清除成功'); // 返回成功消息
    } catch (error) {
        // 避免重复发送响应
        if (!res.headersSent) {
            res.status(500).send('伺服器錯誤');
        }
    }
});



// 創建 HTTP 端點和 Socket.IO 伺服器
const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: '*', // 確保允許来自特定源的請求
    methods: ['GET', 'POST','PUT'],
  },
});

// Socket.IO 連接管理
let onlineUsers = 0;  // 計數線上人數
io.on('connection', (socket) => {
  console.log('使用者上線。');

  // 當用戶加入房間時
  socket.on('joinStoreRoom', (storeName) => {
    socket.join(storeName); // socket.join 是用於讓用戶加入房間
    console.log(`使用者加入商店房間：${storeName}`);
    
    // 您現在可以根據需要廣播消息到這個房間
    // 比如廣播當前線上人數
    const onlineUsers = io.sockets.adapter.rooms.get(storeName)?.size || 0; // 獲取如今庫房中的用户數量
    socket.to(storeName).emit('updateUserCount', onlineUsers); // 向其他在此房間的用戶發送當前人數
  });

  socket.on('disconnect', () => {
    console.log('使用者離線。');
  });
});

// 起動伺服器
const PORT = process.env.PORT || 4000
server.listen(PORT, () => {
  console.log(`伺服器正在端口 ${PORT} 上運行`);
});
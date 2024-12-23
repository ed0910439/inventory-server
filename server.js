//server.js
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const http = require('http');
const { Server } = require('socket.io');
const ExcelJS = require('exceljs'); // 確保這行代碼在文件的頂部
const axios = require('axios'); // 加入這一行以引入 axios
const { load } = require('cheerio');
const bodyParser = require('body-parser');
const cheerio = require('cheerio'); // 导入 cheerio
const { exec } = require('child_process');

const multer = require('multer'); // 導入 multer 中間件
const rateLimit = require('express-rate-limit'); // 導入 express-rate-limit 中間件

// 初始化 Express 應用
const app = express();
app.use(cors());
app.use(express.json());
app.use(bodyParser.json());

// 連接到 MongoDB
require('dotenv').config(); // 載入 .env 文件
mongoose.connect(`mongodb+srv://${process.env.MONGO_URL}`, {
  ssl: true,
});

// 定義產品模型
// 初始化 Express 應用後
const productSchema = new mongoose.Schema({
    商品編號: { type: String, required: true },
    商品名稱: { type: String, required: false },
    規格: { type: String, required: false },
    數量: { type: String, required: false },
    單位: { type: String, required: true },
    到期日: { type: String, required: false },
    廠商: { type: String, required: false },
    庫別: { type: String, required: false }, // 更正名稱為庫別
    盤點日期: { type: String, required: false },
    期初庫存: { type: String, required: false }, // 新增欄位：期初庫存

});
// 动态生成集合名称
const currentDate = new Date();
const year = currentDate.getFullYear();
const latesrmonth = String(currentDate.getMonth()).padStart(2, '0');
const month = String(currentDate.getMonth() + 1).padStart(2, '0'); // 注意：月份从0开始，因此需要加1
const day = currentDate.getDate();

// 根據日期決定使用的月份
if (day < 16) {
    month -= 1; // 回到上個月
    if (month === 0) {
        month = 12; // 回到前一年的12月
        year -= 1;
    }
}

app.get('/api/startInventory/:storeName', async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 获取 URL 中的 storeName

    try {
        if (storeName === 'notStart'){
            res.status(204).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因为请求参数有误
        } else {

            const today = `${year}-${month}-${day}`;
            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const latesCollectionName = `${year}${latesrmonth}${storeName}`; // 动态生成集合名称
            const Product = mongoose.model(collectionName, productSchema);

            // 抓取第一份 HTML 新資料
            const firstResponse = await axios.get(`${process.env.HTML_RESPONSE_URL}${today}%27`);
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
                        newProducts.push(product); // 只保存有效的产品
                    }
                }
            });

            // 获取源集合数据进行比对
            const sourceCollection = mongoose.connection.collection(latesCollectionName);
            const inventoryData = await sourceCollection.find({}).toArray(); // 获取源集合数据

            // 处理最新的盘点数据
            const refinedData = inventoryData.map(item => ({
                商品編號: item.商品編號,
                商品名稱: item.商品名稱,
                規格: item.規格 || '',
                數量: '', // 将数量设置为空
                單位: item.單位 || '',
                到期日: '', // 将到期日设置为空
                廠商: item.廠商 || '',
                庫別: item.庫別 || '',
                盤點日期: '', // 将盘点日期设置为空
                期初庫存: item.數量 || '' // 将数量拷贝到期初库存
            }));
            if (refinedData.length > 0) {
                // 將完成的產品信息存入資料庫
                await Product.insertMany(refinedData);
            }

            // 创建一个映射，方便通过商品编号查找
            const inventoryMap = {};
            inventoryData.forEach(item => {
                inventoryMap[item.商品編號] = {
                    庫別: item.庫別 || '待設定', // 如果没有则标记为待设置
                    廠商: item.廠商 || '',
                    期初庫存: item.數量 || '無紀錄' // 将数量重命名为期初库存
                };
            });

            // 更新新产品数据
            const updatedProducts = newProducts.map(product => {
                const sourceData = inventoryMap[product.商品編號]; // 通过商品编号获取对应数据
                if (sourceData) {
                    // 填入库别
                    product.庫別 = sourceData.庫別;
                    product.廠商 = sourceData.廠商;
                    product.期初庫存 = sourceData.期初庫存; // 将数量字段重命名为期初库存
                } else {
                    // 如果没有找到匹配的商品编号，设置库别为待设置
                    product.庫別 = '待設定';
                }
                return product; // 返回更新后的产品对象
            });

            // 从第二个 HTML 数据源抓取数据
            const secondResponse = await axios.get('https://epos.kingza.com.tw:8090/hyisoft.lost/exportpand.aspx?t=panDianItemCS&id=3148&ClassStore_fCheckSetID=');
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
                        secondInventoryData.push(product); // 将有效的产品添加到列表中
                    }
                }
            });

            // 创建一个映射以比对第二个数据源
            const secondInventoryMap = {};
            secondInventoryData.forEach(item => {
                secondInventoryMap[item.商品編號] = item.單位; // 将单位与商品编号映射
            });

            // 更新产品数据，结合第二个数据源中的单位
            updatedProducts.forEach(product => {
                if (secondInventoryMap[product.商品編號]) {
                    product.單位 = secondInventoryMap[product.商品編號]; // 根据商品编号更新单位
                }
            });

            // 返回所有库别为“待設定”的新品项，等待用户填写
            const pendingProducts = updatedProducts.filter(product => product.庫別 === '待設定');

            if (pendingProducts.length > 0) {
                return res.json(pendingProducts); // 返回待用户填写的产品信息
            } else {
                console.log('没有待设置的产品项');
                return res.status(200).json({ message: '没有待设置的产品项' });
            }
        }

    } catch (error) {
        console.error('处理开始盘点请求时出错:', error);
        if (!res.headersSent) {
            return res.status(500).json({ message: '处理请求时出错', error: error.message });
        }
    }
});
// API 端點：保存補齊的新品
app.post('/api/saveCompletedProducts/:storeName', async (req, res) => {

    const storeName = req.params.storeName || 'notStart'; // 获取 URL 中的 storeName

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

    client.connect(443, 'hass.edc-pws.com', () => {
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
// API端點: 獲取期初庫存數據
app.get('/api/version', (req, res) => {
  fs.readFile(versionFilePath, 'utf8', (err, data) => {
    if (err) {
      console.error("讀取文件時出錯:", err);
      return res.status(500).json({ message: '伺服器錯誤' });
    }
    res.json(JSON.parse(data)); // 返回JSON數據
  });
});

const initialStockFilePath = path.join(__dirname, 'archive', '2024_09.json');

// API端點: 獲取期初庫存數據
app.get('/archive/originaldata', (req, res) => {
  fs.readFile(initialStockFilePath, 'utf8', (err, data) => {
    if (err) {
      console.error("讀取文件時出錯:", err);
      return res.status(500).json({ message: '伺服器錯誤' });
    }
    res.json(JSON.parse(data)); // 返回JSON數據
  });
});



// 獲取產品數據的 API
app.get(`/api/products/:storeName`, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 获取 URL 中的 storeName

    try {
        if (storeName === 'notStart') {
            res.status(400).send('門市錯誤'); // 使用 400 Bad Request 返回错误，因为请求参数有误
        } else {

            const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱
            const Product = mongoose.model(collectionName, productSchema);
            const products = await Product.find(); // 獲取產品數據

            // 返回產品數據
            res.json(products);
            res.status(200); // 使用 400 Bad Request 返回错误，因为请求参数有误

        }
    } catch (error) {
            console.error("獲取產品時出錯:", error);
            res.status(500).send('伺服器錯誤');
        }
    
});
// 更新產品数量的 API 端點
app.put('/api/products/:storeName/:productCode/quantity', async (req, res) => {


    const storeName = req.params.storeName; // 獲取 URL 中的 storeName
    const collectionName = `${year}${month}${storeName}`; // 根據年份、月份和門市生成集合名稱

  try {
      const { productCode } = req.params;
      const { 數量 } = req.body;

      // 更新指定產品的数量
      const updatedProduct = await Product.findOneAndUpdate(
          { 商品編號: productCode },
          { 數量: 數量 },
          { new: true }
      );

      if (!updatedProduct) {
          return res.status(404).send('產品未找到');
      }
      // 廣播更新消息给所有用戶
      io.emit('productUpdated', updatedProduct);

      res.json(updatedProduct);
  } catch (error) {
      console.error('更新產品時出錯:', error);
      res.status(400).send('更新失敗');
  }
});

// 更新產品到期日的 API 端點
app.put('/api/products/:storeName/:productCode/expiryDate', async (req, res) => {
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
      // 廣播更新消息给所有用戶
      io.emit('productUpdated', updatedProduct);
      
      res.json(updatedProduct);
  } catch (error) {
      console.error('更新到期日時出錯:', error);
      res.status(400).send('更新失敗');
  }
});


// 設定 rate limiter: 每分鐘最多 5 次請求
const archiveLimiter = rateLimit({
    windowMs: 1 * 60 * 1000, // 1 minute
    max: 5, // limit each IP to 5 requests per windowMs
});

// API 端點處理盤點歸檔請求
app.post('/api/archive/:storeName', archiveLimiter, async (req, res) => {
    const { year, month, password } = req.body;

    // 輸入驗證
    if (!year || !month || !password) {
        return res.status(400).send('年份、月份和密碼是必需的');
    }

    const adminPassword = process.env.PASSWORD; // 
    if (password !== adminPassword) {
        return res.status(403).send('管理員密碼錯誤');
    }

    try {
        // 獲取當前的庫存數據
        const products = await mongoose.model(collectionName).find();

        // 將數據保存到文件中
        const archiveDir = path.join(__dirname, 'archive');
        const filePath = path.resolve(archiveDir, `${year}年${month}月盤`);
        if (!filePath.startsWith(archiveDir)) {
            return res.status(403).send('無效的文件路徑');
        }
        fs.writeFileSync(filePath, JSON.stringify(products, null, 2), 'utf-8');

        // 將數據從資料庫中清除
        await Product.deleteMany();

        res.status(200).send('數據歸檔成功');

    } catch (error) {
        console.error('處理歸檔請求時出錯:', error);
        res.status(500).send('伺服器錯誤');
    }

});




// 創建 HTTP 端點和 Socket.IO 伺服器
const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: '*', // 確保允許来自特定源的請求
    methods: ['GET', 'POST'],
  },
});

// Socket.IO 連接管理
let onlineUsers = 0;  // 計數線上人數

io.on('connection', (socket) => {
  onlineUsers++;
  console.log('使用者上線。 線上人數：' + onlineUsers + '人');
  
  // 發送當前線上人數給所有用戶
  io.emit('updateUserCount', onlineUsers);

  socket.on('disconnect', () => {
    onlineUsers--;
    console.log('使用者離線。 線上人數：' + onlineUsers + '人');
    // 發送更新的線上人數
    io.emit('updateUserCount', onlineUsers);  });

});

// 起動伺服器
const PORT = process.env.PORT || 4000;
server.listen(PORT, () => {
  console.log(`伺服器正在端口 ${PORT} 上運行`);
});
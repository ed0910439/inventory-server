//server.js
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const http = require('http');
const { Server } = require('socket.io');
const ExcelJS = require('exceljs'); // 確保這行代碼在文件的頂部

const multer = require('multer'); // 導入 multer 中間件

// 初始化 Express 應用
const app = express();
app.use(cors());
app.use(express.json());

// 連接到 MongoDB
require('dotenv').config(); // 載入 .env 文件
mongoose.connect(`mongodb+srv://${process.env.MONGO_URI}`, {
  ssl: true,
});

// 定義產品模型
const productSchema = new mongoose.Schema({
  商品編號: { type: String, required: true },
  商品名稱: { type: String, required: true },
  規格: { type: String, required: false },
  數量: { type: Number, required: true },
  單位: { type: String, required: true },
  到期日: { type: Date },
  廠商: { type: String, required: false },
  溫層: { type: String, required: false },
  盤點日期: { type: String, required: false },

});

const Product = mongoose.model('2024年11月_新店京站', productSchema);

// 從 JSON 文件中加載數據到資料庫
fs.readFile(path.join(__dirname, 'inventorydb.products.json'), 'utf-8', async (err, data) => {
  if (err) {
    console.error('讀取 JSON 文件時出錯:', err);
    return;
  }
  try {
      const products = JSON.parse(data);
      
      const insertPromises = products.map(async (product) => {
          const expiryDate = product.到期日?.$date ? new Date(product.到期日.$date) : null;
          const id = product._id?.$oid ? new mongoose.Types.ObjectId(product._id.$oid) : new mongoose.Types.ObjectId();
          
          // 嘗試查找是否存在該產品
          const existingProduct = await Product.findOne({ 商品編號 : product.商品編號 });
      
          if (!existingProduct) {
              // 如果不存在，則建立新的產品資料
              const NewProductModel = mongoose.model('2024年11月_新店京站');
              const newProduct = new NewProductModel({

                  _id: id,
                  商品編號: product.商品編號,
                  商品名稱: product.商品名稱,
                  規格: product.規格 || '',
                  數量: product.數量 || 0,
                  單位: product.單位,
                  到期日: expiryDate,
				  廠商: product.廠商 || '',
				  溫層: product.溫層 || '',
				  盤點日期: product.盤點日期 || '',

              });
              return newProduct.save(); // 保存到資料庫
          } else {
              console.log(`產品 ${product.商品編號} 已存在，跳过插入。`);
          }
      });

      await Promise.all(insertPromises); // 等待所有異步操作完成
      console.log('產品成功加載到資料庫中');
  } catch (error) {
      console.error('處理 JSON 數據時出錯:', error);
  }
});

// 設置JSON文件的路徑
const versionFilePath = path.join(__dirname, 'version.json');

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



// API 端點獲取產品數據
app.get('/api/products', async (req, res) => {
  try {
      const products = await mongoose.model('2024年11月_新店京站').find();
      res.json(products);
  } catch (error) {
      console.error("獲取產品時出錯:", error);
      res.status(500).send('伺服器錯誤');
  }
});

// 更新產品数量的 API 端點
app.put('/api/products/:productCode/quantity', async (req, res) => {
  try {
      const { productCode } = req.params;
      const { 數量 } = req.body;

      // 更新指定產品的数量
      const updatedProduct = await Product.findOneAndUpdate(
          { 商品編號: productCode },
          { 數量: { $eq: 數量 } },
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
app.put('/api/products/:productCode/expiryDate', async (req, res) => {
  try {
      const { productCode } = req.params;
      const { 到期日 } = req.body;

      // 更新指定產品的到期日
      const updatedProduct = await Product.findOneAndUpdate(
          { 商品編號: productCode },
          { 到期日: new Date(到期日) },
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

// 新增產品的 API 端點
app.post('/api/products', async (req, res) => {
  const { 商品編號, 商品名稱, 規格, 數量, 單位, 到期日, 廠商, 溫層, 盤點日期  } = req.body;

  // 輸入驗證
  if (!商品編號 || !商品名稱 || !數量 || !單位) {
      return res.status(400).send('商品編號、商品名稱、數量和單位是必需的');
  }

  try {
      const NewProductModel = mongoose.model('2024年11月_新店京站');
      const newProduct = new NewProductModel({
          商品編號,
          商品名稱,
          規格: 規格 || '',
          數量: 數量 || 0,
          單位,
          到期日: 到期日 ? new Date(到期日) : null,
		  廠商: product.廠商 || '',
		  溫層: product.溫層 || '',
		  盤點日期: product.盤點日期 || '',
      });

      const savedProduct = await newProduct.save(); // 保存到資料庫
      io.emit('productUpdated', savedProduct); // 廣播產品更新消息
      res.status(201).json(savedProduct); // 返回新建立的產品
  } catch (error) {
      console.error('新增產品時出錯:', error);
      res.status(400).
      res.status(400).send('新增產品失敗');
  }
});
// API 端點處理盤點歸檔請求
app.post('/api/archive', async (req, res) => {
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
        const products = await mongoose.model('2024年11月_新店京站').find();

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

// 使用 multer 設定檔案上傳
const upload = multer({
    storage: multer.diskStorage({
        destination: (req, file, cb) => {
            cb(null, 'uploads/'); // 設定檔案儲存目錄
        },
        filename: (req, file, cb) => {
            cb(null, Date.now() + '-' + file.originalname); // 使用時間戳記和原始檔名產生唯一檔名
        },
    }),
    fileFilter: (req, file, cb) => {
        // 只允許上傳 .xls 和 .xlsx 檔案
        if (file.mimetype === 'application/vnd.ms-excel' || file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            cb(null, true);
        } else {
            cb(new Error('只允許上傳 .xls 或 .xlsx 檔案'));
        }
    },
});


// 新增 API 端點：處理開始盤點請求，包含上傳盤點模板和期初數據
app.post('/api/startInventory', upload.fields([{ name: 'inventoryTemplate', maxCount: 1 }, { name: 'initialStockData', maxCount: 1 }]), async (req, res) => {
    try {
        const inventoryTemplate = req.files.inventoryTemplate[0].path;
        const initialStockData = req.files.initialStockData[0].path;

        // 使用 exceljs 讀取檔案
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inventoryTemplate);
        const worksheet = workbook.worksheets[0];

        const inventoryData = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // 跳過標題列
                const product = {
                    商品編號: row.getCell(2).value,
                    商品名稱: row.getCell(3).value,
                    單位: row.getCell(4).value,
                    廠商: row.getCell(5).value,
                    盤點日期: row.getCell(6).value,
                    到期日: row.getCell(7).value ? new Date(row.getCell(7).value) : null, // 將到期日轉換為 Date 物件
                    溫層: row.getCell(8).value,
                    數量: row.getCell(9).value,
                };
                inventoryData.push(product);
            }
        });
// 更新資料庫 - 使用 findAndUpdate 來更新或新增產品，避免資料遺失
        const updatePromises = inventoryData.map(async (product) => {
            const { 商品編號, ...rest } = product; // 將商品編號分開
            const updateResult = await Product.findOneAndUpdate(
                { 商品編號 }, // 根據商品編號查詢
                { $set: rest }, // 更新其他欄位
                { upsert: true, new: true } // upsert: true 表示如果找不到則新增，new: true 表示返回更新後的資料
            );
            // optionally, emit a socket event here to update the client-side
            // io.emit('productUpdated', updateResult);  //記得更新此部分
        });

        await Promise.all(updatePromises); // 等待所有更新完成

        res.json({ message: '盤點數據已成功上傳' });

    } catch (error) {
        console.error('處理開始盤點請求時出錯:', error);
        res.status(500).json({ error: '伺服器錯誤' });
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

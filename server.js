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
const XLSX = require('xlsx');
const e = require('cors');
const { type } = require('os');
const { error } = require('console');
const dayjs = require('dayjs');
// 進度追蹤用 Map
const uploadTasks = new Map(); // key: taskId, value: { percent, done, message }
const { v4: uuidv4 } = require('uuid');


const app = express();
const upload = multer({ storage: multer.memoryStorage() }); // 使用內存存儲，方便直接獲取buffer

require('dotenv').config();


// 中介配置
//app.use(morgan('combined')); // 使用 morgan 記錄 HTTP 請求
app.use(cookieParser());
app.use(cors());
app.use(bodyParser.json({ limit: '10mb' })); // 更新這裡
app.use(bodyParser.urlencoded({ limit: '10mb', extended: true })); // 保持原狀
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
    停用: { type: Boolean, required: false },
    品號: { type: String, required: true },
    廠商: { type: String, required: false },
    庫別: { type: String, required: false },
    品名: { type: String, required: false },
    規格: { type: String, required: false },
    盤點單位: { type: String, required: false },
    本月報價: { type: Number, default: 0 },
    保存期限: { type: String, required: false },
    進貨單位: { type: String, required: false },
    本月進貨: { type: Number, default: 0 },
    期初盤點: { type: Number, default: 0 },
    期末盤點: { type: Number, default: 0 },
    調出: { type: Number, default: 0 },
    調入: { type: Number, default: 0 },
    本月使用量: { type: Number, default: 0 },
    本月食材成本: { type: Number, default: 0 },
    本月萬元用量: { type: Number, default: 0 },
    週用量: { type: Number, default: 0 },
    盤點日期: { type: String, required: false },
    進貨上傳: { type: Boolean, default: false }, // 新增進貨上傳欄位
    盤點完成: { type: Boolean, default: false }, // 新增盤點完成欄位
    最後更新時間: { type: Date, default: Date.now }, // 新增最後更新時間欄位
    最後更新欄位: { type: String, default: '初始化設定' } // 新增最後更新欄位
});

// 定義 sanitizeInput 函數
const sanitizeInput = (input) => {
    return encodeURIComponent(input.trim());
};
// 動態產生集合品名
// --- 核心日期邏輯計算區塊 (只需要放在程式碼最前端) ---

const currentDate = new Date();
const currentDay = currentDate.getDate();

// 1. 計算「基準月份 (current/target)」的 Date 物件
const targetDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);

// 判斷：如果日期在 1 號到 15 號之間，基準月份應為上個月
if (currentDay <= 15) {
    // 往前推一個月，Date 物件會自動處理跨年
    targetDate.setMonth(targetDate.getMonth() - 1);
}

// 2. 計算「上一個月份 (last/previous target)」的 Date 物件
const lastTargetDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
lastTargetDate.setMonth(lastTargetDate.getMonth() - 1);


// --- 將計算結果賦值給您的舊變數 (確保後續程式碼無需改動) ---

// 基準月份 (您的 old year/month)
let year = targetDate.getFullYear();
let month = targetDate.getMonth() + 1; // 1-12 格式

// 上一個基準月份 (您的 old lastYear/lastMonth)
let lastYear = lastTargetDate.getFullYear();
let lastMonth = lastTargetDate.getMonth() + 1; // 1-12 格式

// 格式化月份為兩位數 (保留您的格式化變數，以滿足舊程式碼的需求)
const formattedMonth = String(month).padStart(2, '0');
const formattedLastMonth = String(lastMonth).padStart(2, '0');


// --- 範例：計算 tdate (可選，如果您想保留) ---

// 取得基準月份的月底日 (利用 setDate(0) 技巧)
const targetMonthIndex = targetDate.getMonth(); // 0-11
const endOfMonthDate = new Date(year, targetMonthIndex + 1, 0).getDate();
const tdate = `${year}-${formattedMonth}-${endOfMonthDate}`;
const startDate = `${year}/${formattedMonth}/01`;
const endDate = `${year}/${formattedMonth}/${endOfMonthDate}`;

console.log(tdate);
console.log(year, formattedMonth, currentDay); // 輸出當前年份、月份、日期
console.log(lastYear, formattedLastMonth, currentDay); // 輸出上月份的年份、月份、日期

const currentlyEditingProductsByStore = {};

// 開始盤點
app.get('/api/startInventory/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart';
    try {
        if (storeName === 'notStart') {
            return res.status(204).send('尚未選擇門市');
        }

        const collectionTmpName = `${year}${formattedMonth}${storeName}_tmp`;
        const collectionName = `${year}${formattedMonth}${storeName}`;
        const lastCollectionName = `${lastYear}${formattedLastMonth}${storeName}`;
        const lastCollectionDemo = `${lastYear}${formattedLastMonth}dc03021`;

        // 動態 model
        const ProductTemp = mongoose.models[collectionTmpName] || mongoose.model(collectionTmpName, productSchema);
        const Product = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);

        // 如果當月正式盤點集合已存在，阻止重複建立
        const Products = await Product.find({});
        if (Products.length > 0) {
            return res.status(500).json({
                message: '當月已有建立盤點單，請依照下方指示操作後，再操作一次！',
                error: '選右上方選單>盤點結束>輸入密碼>勾選注意事項>盤點清除'
            });
        }

        await mongoose.connection.collection(collectionTmpName).deleteMany({});

        // --- 從新 API 抓取數據 ---
        const apiUrl = "https://kingzaap.unium.com.tw/BohAPI/MSCINKX/FindInventoryData";
        const payload = { Str_No: storeName === 'dc03021test' ? 'dc03021' : storeName, Tdate: tdate, BrandNo: "004" };
        console.log(payload);
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) throw new Error(`API 回傳錯誤: ${response.status}`);
        const result = await response.json();

        if (result.returnCode !== "200" || !Array.isArray(result.data) || result.data.length === 0) {
            return res.status(404).json({ message: 'API 沒有回傳有效盤點資料' });
        }

        // 將 API 資料轉換成乾淨格式
        const apiProducts = [];
        if (result.returnCode === "200" && result.data?.length > 0) {
            result.data.forEach(item => {
                apiProducts.push({
                    品號: item.Goo_No,
                    品名: item.Goo_Na,
                    規格: item.Memo || '',
                    盤點單位: item.Unit || ''
                });
            });
        }

        // --- 取得上個月盤點資料（無資料套用 DEMO） ---
        let lastMonthInventory = [];
        const lastMonthInventoryMap = new Map();

        const collections = await mongoose.connection.db.listCollections({ name: lastCollectionName }).toArray();
        if (collections.length > 0) {
            // 上個月有正式 collection
            lastMonthInventory = await mongoose.connection.collection(lastCollectionName).find({}).toArray();
        } else {
            // 上個月沒有資料，套用 DEMO 模板，期初盤點設為 0
            lastMonthInventory = await mongoose.connection.collection(lastCollectionDemo).find({}).toArray();
            lastMonthInventory = lastMonthInventory.map(item => ({ ...item, 期初盤點: 0 }));
        }

        // 轉為 Map 以便快速查詢
        lastMonthInventory.forEach(item => lastMonthInventoryMap.set(item.品號, item));

        // --- 資料合併邏輯 ---
        const productsNeedingSetup = [];          // 前端需要設定的品項
        const allProductsToInsertIntoTemp = [];   // 最終插入暫存集合的品項

        apiProducts.forEach(apiProduct => {
            // 僅處理 API 有回傳的品項
            let currentProduct = { ...apiProduct };
            const lastMonthProduct = lastMonthInventoryMap.get(apiProduct.品號);

            if (lastMonthProduct) {
                // 上月有 → 繼承相關資訊
                currentProduct.庫別 = lastMonthProduct.庫別;
                currentProduct.廠商 = lastMonthProduct.廠商;
                currentProduct.停用 = lastMonthProduct.停用 || false;
                currentProduct.期初盤點 = lastMonthProduct.期末盤點 || 0;
                currentProduct.期末盤點 = 0;
                currentProduct.盤點日期 = tdate;
                currentProduct.本月報價 = 0;
                currentProduct.保存期限 = '';
                currentProduct.進貨單位 = lastMonthProduct.進貨單位 || '';
                currentProduct.本月進貨 = 0;
                currentProduct.調出 = 0;
                currentProduct.調入 = 0;
                currentProduct.本月使用量 = 0;
                currentProduct.本月食材成本 = 0;
                currentProduct.本月萬元用量 = 0;
                currentProduct.週用量 = 0;
                currentProduct.盤點完成 = false;
                currentProduct.進貨上傳 = false;
                currentProduct.最後更新時間 = new Date();
                currentProduct.最後更新欄位 = '初始化設定';
                currentProduct.規格 = currentProduct.規格 || lastMonthProduct.規格 || '';
                currentProduct.盤點單位 = currentProduct.盤點單位 || lastMonthProduct.盤點單位 || '';
                currentProduct.品名 = currentProduct.品名 || lastMonthProduct.品名 || '';

            } else {
                // 上月沒有 → 視為新商品
                currentProduct.庫別 = '待設定';
                currentProduct.廠商 = '待設定';
                currentProduct.停用 = false;
                currentProduct.期初盤點 = 0;
                currentProduct.期末盤點 = 0;
                currentProduct.盤點日期 = tdate;
                currentProduct.本月報價 = 0;
                currentProduct.保存期限 = '';
                currentProduct.進貨單位 = '';
                currentProduct.本月進貨 = 0;
                currentProduct.調出 = 0;
                currentProduct.調入 = 0;
                currentProduct.本月使用量 = 0;
                currentProduct.本月食材成本 = 0;
                currentProduct.本月萬元用量 = 0;
                currentProduct.週用量 = 0;
                currentProduct.盤點完成 = false;
                currentProduct.進貨上傳 = false;
                currentProduct.最後更新時間 = new Date();
                currentProduct.最後更新欄位 = '初始化設定';
                currentProduct.規格 = currentProduct.規格 || '';
                currentProduct.盤點單位 = currentProduct.盤點單位 || '';
                currentProduct.品名 = currentProduct.品名 || '';

                productsNeedingSetup.push(currentProduct);
            }

            allProductsToInsertIntoTemp.push(currentProduct);
        });

        // ❗ 注意：如果 lastMonthInventory 有，但 API 沒有 → 自動略過，不加進本月資料
        // → 因為我們只依據 API 有回傳的品號進行建立

        // --- 批量插入暫存集合 ---
        if (allProductsToInsertIntoTemp.length > 0) {
            await ProductTemp.insertMany(allProductsToInsertIntoTemp);
        }

        // --- 決策回傳 ---
        if (productsNeedingSetup.length > 0) {
            // 有新商品 → 需前端補設定
            return res.json(productsNeedingSetup);
        } else {
            // 全部商品自動完成 → 寫入正式集合
            await Product.insertMany(allProductsToInsertIntoTemp);
            const tempExists = await mongoose.connection.db.listCollections({ name: collectionTmpName }).toArray();
            if (tempExists.length > 0) {
                await ProductTemp.collection.drop();
            }
            return res.status(201).json({ message: '所有產品已自動保存並更新正式庫存' });
        }

    } catch (error) {
        console.error('建立盤點資料庫時發生錯誤:', error);
        return res.status(500).json({ message: '建立盤點資料庫失敗', error: error.message });
    }
});

// 保存已完成產品的 API 端點
app.post('/api/saveCompletedProducts/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart';
    const collectionName = `${year}${formattedMonth}${storeName}`;
    const collectionTmpName = `${year}${formattedMonth}${storeName}_tmp`;

    const Product =
        mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);
    const ProductTemp =
        mongoose.models[collectionTmpName] || mongoose.model(collectionTmpName, productSchema);

    const message = '請將頁面重新整理以更新品項';
    const inventoryItems = req.body; // 需更新庫別/廠商/停用

    try {
        if (!Array.isArray(inventoryItems) || inventoryItems.length === 0) {
            return res.status(400).json({ message: '請求體中沒有有效的產品數據' });
        }

        // --- 更新暫存集合中的新品設定 ---
        const updateResults = await Promise.all(
            inventoryItems.map((item) => {
                if (!item || !item.品號 || item.庫別 === undefined || item.廠商 === undefined) {
                    console.warn(`缺少必要欄位，跳過更新品號: ${item?.品號}`);
                    return Promise.resolve({ matchedCount: 0, modifiedCount: 0 });
                }
                return ProductTemp.updateOne(
                    { 品號: item.品號 },
                    { $set: { 庫別: item.庫別, 廠商: item.廠商, 停用: !!item.停用 } }
                );
            })
        );

        const totalModified = updateResults.reduce((sum, r) => sum + (r.modifiedCount || 0), 0);
        console.log(`成功更新 ${totalModified} 個產品`);

        // --- 取出暫存集合所有品項 ---
        let completedTmpProducts = await ProductTemp.find();

        // ✅ 去重：依品號保留最後一筆
        const uniqueMap = new Map();
        for (const item of completedTmpProducts) {
            uniqueMap.set(item.品號, item);
        }
        completedTmpProducts = Array.from(uniqueMap.values());

        // ✅ 升序排列：依品號排序
        completedTmpProducts.sort((a, b) => String(a.品號).localeCompare(String(b.品號)));

        // --- 寫入正式集合 ---
        if (completedTmpProducts.length > 0) {
            await Product.deleteMany({});
            await Product.insertMany(completedTmpProducts);

            // 清空暫存集合
            try {
                await ProductTemp.collection.drop();
            } catch (err) {
                if (err.code !== 26) console.error('刪除暫存集合失敗:', err);
            }

            return res.status(201).json({
                message: '所有新產品已成功保存並更新正式庫存（已去重並排序）'
            });
        } else {
            return res.status(200).json({ message: '沒有需要保存的新產品' });
        }

    } catch (error) {
        console.error('儲存產品時出錯:', error);
        return res.status(500).json({
            message: '儲存失敗',
            error: process.env.NODE_ENV === 'development' ? error.stack : error.message
        });
    }
});

// --- 新增 API 端點: 同步當月盤點量 (回傳 API 資料 + 整合本地期末盤點) ---
app.get('/api/syncInventoryData/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart';
    const collectionName = `${year}${formattedMonth}${storeName}`; // 當月正式集合

    if (storeName === 'notStart') {
        return res.status(400).json({ message: '門市名稱不可為空' });
    }

    try {
        const currentTdate = tdate;
        const Product = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);

        // 1. 取得本地資料庫中的所有品項，轉成 Map 以便快速查詢
        const localProducts = await Product.find({}).select('品號 期末盤點');
        const localInventoryMap = new Map();
        localProducts.forEach(p => {
            // 存儲 品號: 期末盤點
            localInventoryMap.set(p.品號, p.期末盤點 || 0);
        });

        // 2. 呼叫外部 API
        const apiUrl = "https://kingzaap.unium.com.tw/BohAPI/MSCINKX/FindInventoryData";
        const apiStrNo = storeName === 'dc03021test' ? 'dc03021' : storeName;
        const payload = { Str_No: apiStrNo, Tdate: currentTdate, BrandNo: "004" };

        const response = await axios.post(apiUrl, payload, {
            headers: { 'Content-Type': 'application/json' }
        });

        const result = response.data;

        if (result.returnCode !== "200" || !Array.isArray(result.data)) {
            console.warn(`[Sync API] 門市 ${storeName} 於 ${currentTdate} 沒有回傳有效盤點資料:`, result.returnMsg);
            return res.json([]);
        }

        // 3. 整合資料：提取並加入本地盤點量
        const processedData = result.data.map(item => {
            const productCode = item.Goo_No;
            const localQty = localInventoryMap.get(productCode); // 取得本地期末盤點量

            return {
                Goo_No: productCode,              // API 品號
                Goo_Na: item.Goo_Na,              // API 品名
                Api_Qty: item.Tto_Qty,            // API 盤點量 (遠端)
                Unit: item.Unit,                  // 單位
                Local_Qty: localQty !== undefined ? localQty : 0 // 本地盤點量 (MongoDB: 期末盤點)
            };
        });
        // 🚨 解決 304 快取問題的關鍵步驟 🚨
        res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
        res.set('Pragma', 'no-cache');
        res.set('Expires', '0');
        res.set('Surrogate-Control', 'no-store');

        // 將轉換後的陣列回傳給前端
        res.json(processedData);

    } catch (error) {
        console.error(`[Sync API] 同步門市 ${storeName} 盤點量時發生錯誤:`, error.message);
        return res.status(500).json({ message: '同步盤點量資料失敗', error: error.message });
    }
});
// 處理匯出至總表的 API 路由 (整合銷售排行與格式保留)
app.post('/api/export-master-sheet/:storeName', upload.single('excelFile'), async (req, res) => {
    
    if (!req.file) {
        return res.status(400).json({ success: false, message: '請上傳 Excel 總表檔案。' });
    }

    const storeName = req.params.storeName || 'notStart';

    if (storeName === 'notStart') {
        return res.status(400).json({ message: '門市名稱不可為空' });
    }
    
    // 🚨 步驟 2: 從 MongoDB 取得數據並呼叫外部 API
    let inventoryData;
    let salesData = []; // 用於儲存 API 返回的銷售數據

    // 2.1 從 MongoDB 抓取數據 (保留原有邏輯)
    try {
        // 假設 productSchema 和 mongoose.model 均已定義
        const collectionName = `${year}${formattedMonth}${storeName}`;
        const InventoryModel = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);
        inventoryData = await InventoryModel.find({ /* filter */ }).lean();

        if (!inventoryData || inventoryData.length === 0) {
            // 如果 MongoDB 數據為空，仍然嘗試抓取銷售數據
            console.warn('MongoDB 未回傳任何可供寫入的庫存數據。');
        }
    } catch (e) {
        console.error("從 MongoDB 抓取數據失敗:", e);
        return res.status(500).json({ success: false, message: '後端數據庫查詢錯誤。' });
    }

    // 2.2 🚨 呼叫外部 API 取得銷售數據
    try {
        const apiUrl = "https://kingzaap.unium.com.tw/SettingApi/KzSale/getGooForExcel";
        const apiBody = {
            kind: "M",
            Str_No: storeName,
            StartDate: startDate,
            EndDate: endDate,
            Op_Type: ""
        };
        const apiHeaders = {
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json",
            "Referer": "https://kingzaap.unium.com.tw/BackWeb/Report/ProductAnalysis"
        };

        const apiResponse = await axios.post(apiUrl, apiBody, { headers: apiHeaders });

        if (apiResponse.data && apiResponse.data.returnCode === "200" && Array.isArray(apiResponse.data.data)) {
            salesData = apiResponse.data.data;
        }
    } catch (apiError) {
        console.warn("外部 API 銷售數據抓取失敗，將跳過寫入 '商品銷售排行(貼)' 工作表:", apiError.message);
    }
    
    // 步驟 3: Excel 讀取、寫入與格式復原
    try {
        const fileBuffer = req.file.buffer;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer); 

        const targetSheetName = '總表';
        const worksheet = workbook.getWorksheet(targetSheetName);

        if (!worksheet) {
            return res.status(400).json({ success: false, message: `Excel 檔案中找不到工作表："${targetSheetName}"` });
        }

        // =========================================================================
        // 🚨 3.1：總表格式快照 (在修改數據前保存)
        // =========================================================================
        
        // 1. 保存合併儲存格範圍 (A1:B1 格式的字串)
        const mergedCellsRanges = [];
        if (worksheet.model.merges) {
            for (const range in worksheet.model.merges) {
                mergedCellsRanges.push(range);
            }
        }
        
        // 2. 保存工作表的格線設定
        const isGridLinesVisible = worksheet.views && worksheet.views.length > 0
            ? worksheet.views[0].showGridLines
            : true; 
            
        // 3. 儲存格底色/樣式快照 (只快照前兩行)
        const cellStylesSnapshot = {}; 
        for (let i = 1; i <= 2; i++) {
            worksheet.getRow(i).eachCell({ includeEmpty: true }, (cell) => {
                const cellRef = cell.address;
                if (cell.fill) {
                    cellStylesSnapshot[cellRef] = { fill: cell.fill };
                }
            });
        }
        
        // 3.2 抬頭定位與檢查 (使用 exceljs 遍歷)
        const headerRow = worksheet.getRow(2); 

        if (!headerRow || headerRow.values.length <= 1) {
            return res.status(400).json({ success: false, message: '無法讀取工作表 "總表" 的抬頭行。' });
        }
        
        const requiredHeaders = [
            '品號', '本月進貨', '期初盤點', '期末盤點', '調出', '調入'
        ];
        
        const headersMap = {};
        
        headerRow.eachCell((cell, colNumber) => {
            const title = String(cell.value || '').trim();
            
            if (requiredHeaders.includes(title)) {
                // 儲存為 0-based 索引
                headersMap[title] = colNumber - 1; 
            }
        });

        for (const key of requiredHeaders) {
            if (headersMap[key] === undefined) { 
                 return res.status(400).json({ success: false, message: `工作表 "總表" 缺少關鍵欄位抬頭："${key}"` });
            }
        }

        // 建立數據庫數據的品號查找表 (Map)
        const dbDataMap = inventoryData.reduce((acc, item) => {
            if (item['品號']) { 
                acc[String(item['品號']).trim()] = item;
            }
            return acc;
        }, {});

        // 3.3 遍歷 Excel 數據行，寫入數據 (只修改 value)
        worksheet.eachRow({ includeEmpty: false, firstRow: 3 }, (row) => {
            const productCodeColIndex = headersMap['品號'] + 1; 
            const excelProductCodeCell = row.getCell(productCodeColIndex);
            const excelProductCode = String(excelProductCodeCell.value || '').trim();
            
            const matchedItem = dbDataMap[excelProductCode]; 

            if (matchedItem) {
                const updateCell = (headerName, value) => {
                    const colIndex = headersMap[headerName] + 1; 
                    if (colIndex !== 0) {
                        const cell = row.getCell(colIndex);
                        // 最終保守寫法：只修改值
                        cell.value = Number(value) || 0; 
                    }
                };
                
                updateCell('本月進貨', matchedItem['本月進貨']);
                updateCell('期初盤點', matchedItem['期初盤點']);
                updateCell('期末盤點', matchedItem['期末盤點']);
                updateCell('調出', matchedItem['調出']);
                updateCell('調入', matchedItem['調入']);
            }
        });

        // =========================================================================
        // 🚨 3.4：總表格式復原 (自動欄寬計算 + 樣式還原)
        // =========================================================================
        
        // 1. 實作自動欄寬計算
        const columnsToAutoFit = ['品號', '本月進貨', '期初盤點', '期末盤點', '調出', '調入'];
        const columnWidths = {}; 

        // 初始化 max width for required columns (使用抬頭長度作為起點)
        columnsToAutoFit.forEach(header => {
            const colIndex = headersMap[header]; 
            if (colIndex !== undefined) {
                columnWidths[colIndex + 1] = String(header).length; 
            }
        });

        // 遍歷所有行，計算內容最大長度
        worksheet.eachRow({ includeEmpty: false }, (row) => {
            columnsToAutoFit.forEach(header => {
                const colIndex = headersMap[header]; 
                const colNum = colIndex + 1; 
                const cell = row.getCell(colNum); 
                
                if (cell.value) {
                    let content = String(cell.value);
                    if (cell.value && typeof cell.value === 'object' && cell.value.text) {
                        content = cell.value.text;
                    }
                    const currentLength = content.length;
                    if (currentLength > (columnWidths[colNum] || 0)) {
                        columnWidths[colNum] = currentLength;
                    }
                }
            });
        });

        // 設置欄寬
        for (const colNum in columnWidths) {
            const length = columnWidths[colNum];
            const newWidth = Math.max(10, length * 1.25); // 最小寬度 10
            try {
                 worksheet.getColumn(parseInt(colNum)).width = newWidth;
            } catch (e) {
                 console.log(`總表自動設定欄寬失敗 (欄 ${colNum}):`, e.message);
            }
        }
        
        // 2. 恢復合併儲存格
        mergedCellsRanges.forEach(range => {
            try {
                worksheet.unmergeCells(range); 
                worksheet.mergeCells(range);   
            } catch (e) {
                console.log(`總表無法重新合併儲存格範圍 ${range}:`, e.message);
            }
        });
        
        // 3. 恢復格線設定
        worksheet.views = [{ 
            state: 'normal', 
            showGridLines: isGridLinesVisible 
        }];

        // 4. 恢復儲存格底色/樣式
        for (const cellRef in cellStylesSnapshot) {
            try {
                const cell = worksheet.getCell(cellRef);
                if (cellStylesSnapshot[cellRef].fill) {
                    cell.fill = cellStylesSnapshot[cellRef].fill;
                }
            } catch (e) {
                console.log(`總表恢復儲存格 ${cellRef} 樣式失敗:`, e.message);
            }
        }

// =========================================================================
        // 🚨 步驟 3.5：建立銷售數據和計算合計
        // =========================================================================
        let totalSalesNet = 0; 
        let salesRows = [];
        const totals = {}; 
        const totalHeaders = [
            '總銷量', '銷售毛額', '銷售淨額', '單品折扣讓', 
            '全單攤抵額', 'PSD', '占比'
        ];
        totalHeaders.forEach(h => totals[h] = 0);

        if (salesData.length > 0) {
            salesRows = salesData.map(item => {
                const rowData = {
                    '分店名稱': item['店代號'] || storeName,
                    '商品條碼': item['商品編號'] || '',
                    '商品名稱': item['商品名稱'] || '',
                    '總銷量': Number(item['總銷量']) || 0,
                    '銷售毛額': Number(item['銷售毛額']) || 0,
                    '銷售淨額': Number(item['銷售淨額']) || 0,
                    '單品折扣讓': Number(item['單品折扣讓']) || 0,
                    '全單攤抵額': Number(item['全單攤抵額']) || 0,
                    'PSD': Number(item['PSD']) || 0,
                    '占比': item['業績佔比'] !== null ? Number(item['業績佔比']) : 0
                };

                // 計算總和
                totalHeaders.forEach(h => { totals[h] += rowData[h]; });
                return rowData;
            });
            totalSalesNet = totals['銷售淨額']; 
        }

        // 🚨 步驟 3.6: 寫入銷售淨額合計值到總表 E1
        const totalSheet = workbook.getWorksheet(targetSheetName); // '總表'
        if (totalSheet) {
            const E1Cell = totalSheet.getCell('E1'); 
            E1Cell.value = totalSalesNet;
            E1Cell.numFmt = '#,##0'; // 設定千分位格式
            E1Cell.font = { bold: true };
        }


        // =========================================================================
        // 🚨 步驟 3.7：建立並寫入 "商品銷售排行(貼)" 工作表 (含合計行)
        // =========================================================================
        if (salesRows.length > 0) {
            const newSheetName = "商品銷售排行(貼)";
            const existingWorksheet = workbook.getWorksheet(newSheetName);

            // 1. 檢查並移除舊工作表
            if (existingWorksheet) {
                workbook.removeWorksheet(existingWorksheet.id); 
                console.log(`已移除既有的工作表: ${newSheetName}，準備重新建立。`);
            }

            const newWorksheet = workbook.addWorksheet(newSheetName);

            const newSheetHeaders = [
                '分店名稱', '商品條碼', '商品名稱', '總銷量', '銷售毛額', 
                '銷售淨額', '單品折扣讓', '全單攤抵額', 'PSD', '占比'
            ];
            
            // 2. 設定抬頭
            newWorksheet.columns = newSheetHeaders.map(header => ({
                header: header,
                key: header,
                width: 10
            }));

            // 3. 寫入數據
            newWorksheet.addRows(salesRows);

            // 4. 🚨 加入合計行
            const totalRowData = {
                '商品名稱': '合計', 
                '總銷量': totals['總銷量'],
                '銷售毛額': totals['銷售毛額'],
                '銷售淨額': totals['銷售淨額'],
                '單品折扣讓': totals['單品折扣讓'],
                '全單攤抵額': totals['全單攤抵額'],
                'PSD': totals['PSD'],
                '占比': totals['占比']
            };
            
            const totalRow = newWorksheet.addRow(totalRowData);

            // 5. 對合計行應用粗體樣式
            totalRow.font = { bold: true };
            
            // 6. 🚨 對合計欄位應用數字格式
            const headerIndices = { 
                '總銷量': 4, '銷售毛額': 5, '銷售淨額': 6, 
                '單品折扣讓': 7, '全單攤抵額': 8, 'PSD': 9, '占比': 10
            };
            
            totalHeaders.forEach(header => {
                const colNum = headerIndices[header];
                if (colNum) {
                    const cell = totalRow.getCell(colNum);
                    if (header === '總銷量' || header === 'PSD') {
                        cell.numFmt = '#,##0.00'; // 兩位小數
                    } else if (header === '占比') {
                        cell.numFmt = '0.00%'; // 百分比格式
                    } else {
                        cell.numFmt = '#,##0'; // 銷售額/折扣等用千分位
                    }
                }
            });

            // 7. 自動欄寬調整
            newWorksheet.columns.forEach(column => {
                let maxContentLength = String(column.header).length;
                newWorksheet.getColumn(column.number).eachCell({ includeEmpty: true }, (cell) => {
                    let content = String(cell.value || '');
                    if (cell.value && typeof cell.value === 'object' && cell.value.text) { content = cell.value.text; }
                    if (content.length > maxContentLength) { maxContentLength = content.length; }
                });
                const newWidth = Math.max(10, maxContentLength * 1.25);
                newWorksheet.getColumn(column.number).width = newWidth;
            });
        }
        
        // 步驟 4: 輸出修改後的檔案並傳送給前端
        const newWorkbookBuffer = await workbook.xlsx.writeBuffer(); 
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=' + encodeURIComponent('更新後的盤點總表.xlsx'));
        res.send(newWorkbookBuffer);

    } catch (error) {
        console.error('後端 Excel 處理發生錯誤:', error);
        res.status(500).json({ success: false, message: error.message || '伺服器處理 Excel 檔案時發生錯誤。' });
    }
});
// --- 新增 API 端點: 批量更新期末盤點量 ---
app.post('/api/batchUpdateInventoryQty/:storeName', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart';
    const collectionName = `${year}${formattedMonth}${storeName}`;
    const Product = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);

    // 接收前端傳來的需要更新的品項陣列 (包含 Goo_No 和 Api_Qty)
    const updateItems = req.body;

    if (storeName === 'notStart') {
        return res.status(400).json({ message: '門市名稱不可為空' });
    }
    if (!Array.isArray(updateItems) || updateItems.length === 0) {
        return res.status(400).json({ message: '請求中沒有有效的更新數據' });
    }

    try {
        const bulkOps = updateItems.map(item => {
            const finalQty = parseFloat(item.Api_Qty) || 0; // 確保是數字

            return {
                updateOne: {
                    filter: { 品號: item.Goo_No }, // 根據品號過濾
                    update: {
                        $set: {
                            期末盤點: finalQty,
                            盤點完成: true,
                            最後更新時間: new Date(),
                            最後更新欄位: "期末盤點(批量)" // 符合您的要求
                        }
                    }
                }
            };
        });

        const result = await Product.bulkWrite(bulkOps);

        // 成功後廣播更新給前端
        io.to(storeName).emit('bulkInventoryUpdateCompleted', {
            message: `成功批量更新 ${result.modifiedCount} 筆盤點量`,
            modifiedCount: result.modifiedCount
        });

        res.status(200).json({
            message: `成功批量更新 ${result.modifiedCount} 筆盤點量`,
            modifiedCount: result.modifiedCount
        });

    } catch (error) {
        console.error(`[Batch Update API] 批量更新盤點量時發生錯誤:`, error.message);
        res.status(500).json({ message: '批量更新盤點量失敗', error: error.message });
    }
});

// 上傳盤點數量到 Kingzaap
app.post('/api/upload-inventory/:storeName', (req, res) => {
    const storeName = req.params.storeName;
    const taskId = uuidv4(); // 生成唯一 taskId
    const collectionName = `${year}${formattedMonth}${storeName}`;
    const Product = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);

    // 先回傳 taskId 給前端
    res.json({ success: true, taskId, message: '上傳任務已啟動' });

    // 非同步上傳任務
    setImmediate(async () => {
        uploadTasks.set(taskId, { percent: 0, done: false, message: '初始化' });
        try {
            const localProducts = await Product.find({});

            // 取得 API 資料
            const apiUrl = "https://kingzaap.unium.com.tw/BohAPI/MSCINKX/FindInventoryData";
            const payload = { Str_No: storeName, Tdate: tdate, BrandNo: "004" };
            const response = await axios.post(apiUrl, payload, {
                headers: { 'Content-Type': 'application/json' }
            });

            if (!Array.isArray(response.data.data)) {
                uploadTasks.set(taskId, { percent: 0, done: true, message: 'FindInventoryData 回傳錯誤' });
                return;
            }

            // 更新 Tto_Qty 與 Unit_Qty
            const updatedData = response.data.data.map(apiItem => {
                const localItem = localProducts.find(p => p.品號 === apiItem.Goo_No);
                const qty = localItem ? Number(localItem.期末盤點 || 0) : 0;
                return {
                    ...apiItem,
                    Tto_Qty: qty,
                    Unit_Qty: qty.toString()
                };
            });

            // 一次性上傳到 UpsertInventoryData
            await axios.post(
                "https://kingzaap.unium.com.tw/BohAPI/MSCINKX/UpsertInventoryData",
                updatedData,
                { headers: { 'Content-Type': 'application/json' } }
            );

            uploadTasks.set(taskId, { percent: 100, done: true, message: '上傳完成' });
        } catch (error) {
            console.error(error);
            uploadTasks.set(taskId, { percent: 0, done: true, message: '上傳失敗' });
        }
    });
});

// 查詢上傳進度
app.get('/api/upload-inventory/:storeName/:taskId', (req, res) => {
    const { taskId } = req.params;
    const task = uploadTasks.get(taskId);
    if (!task) return res.status(404).json({ message: '找不到任務' });
    res.json(task);
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

// 抓取本月進貨量並更新資料庫
app.post('/api/fetchMonthlyPurchase/:storeName', async (req, res) => {
    const storeName = req.params.storeName;

    try {
        // 假設 year 和 formattedMonth 已在後端計算或可透過函數取得
        const lastDay = new Date(year, formattedMonth, 0).getDate(); // 自動取當月最後一天

        const startDate = `${year}-${String(formattedMonth).padStart(2, '0')}-01`;
        const endDate = `${year}-${String(formattedMonth).padStart(2, '0')}-${lastDay}`;

        // 抓取 Kingzaap API
        const apiResponse = await axios.post(
            "https://kingzaap.unium.com.tw/BohAPI/MSCPURX/GetAcceptanceMaterialsQuery",
            {
                Str_No: storeName,
                Start_Time: startDate,
                End_Time: endDate
            },
            {
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/plain, */*"
                }
            }
        );

        const data = apiResponse.data?.data || [];
        if (!Array.isArray(data)) return res.status(500).json({ message: "API 回傳格式錯誤" });

        // 將 Total_Qty 轉成 { 品號: 總數量 } 映射
        const totalQtyMap = {};
        data.forEach(item => {
            const productCode = item.Goo_No?.trim();
            if (!productCode) return;
            totalQtyMap[productCode] = Number(item.Total_Qty || 0);
        });

        // 更新資料庫
        const collectionName = `${year}${String(month).padStart(2, '0')}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);

        const bulkOps = [];
        for (const [productCode, qty] of Object.entries(totalQtyMap)) {
            const roundedQty = parseFloat(qty.toFixed(2));
            bulkOps.push({
                updateOne: {
                    filter: { 品號: productCode },
                    update: { $set: { 本月進貨: roundedQty, 進貨上傳: true } }

                }
            });
        }

        if (bulkOps.length > 0) {
            await Product.bulkWrite(bulkOps);
            return res.json({ message: `成功更新 ${bulkOps.length} 筆本月進貨資料` });
        } else {
            return res.json({ message: "無資料需要更新" });
        }

    } catch (err) {
        console.error(err);
        return res.status(500).json({ message: "抓取或更新本月進貨失敗", error: err.message });
    }
});

// 抓取調入資料並更新資料庫
app.post('/api/fetchCallUpData/:storeName', async (req, res) => {
    const storeName = req.params.storeName;

    try {
        // 後端計算特殊系統年月
        const lastDay = new Date(year, formattedMonth, 0).getDate(); // 自動取當月最後一天

        const startDate = `${year}-${String(formattedMonth).padStart(2, '0')}-01`;
        const endDate = `${year}-${String(formattedMonth).padStart(2, '0')}-${lastDay}`;

        // 抓取 Kingzaap 調入資料
        const apiResponse = await axios.post(
            "https://kingzaap.unium.com.tw/BohAPI/MSCTTOMI/FindCallUpData",
            {
                Des_StrNo: storeName,
                StartTime: startDate,
                EndTime: endDate
            },
            {
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/plain, */*"
                }
            }
        );

        const data = apiResponse.data?.data || [];
        if (!Array.isArray(data)) return res.status(500).json({ message: "API 回傳格式錯誤" });

        // 生成品號 -> 調入數量映射
        const callUpMap = {};
        data.forEach(item => {
            const productCode = item.Goo_No?.trim();
            if (!productCode) return;

            // 累加同一品號的數量
            if (callUpMap[productCode]) {
                callUpMap[productCode] += Number(item.Qty || 0);
            } else {
                callUpMap[productCode] = Number(item.Qty || 0);
            }
        });

        // 更新 MongoDB
        const collectionName = `${year}${String(month).padStart(2, '0')}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);

        const bulkOps = [];
        for (const [productCode, qty] of Object.entries(callUpMap)) {
            const roundedFinalQty = parseFloat(qty.toFixed(2));
            bulkOps.push({
                updateOne: {
                    filter: { 品號: productCode },
                    update: { $set: { 調入: roundedFinalQty } }
                }
            });
        }

        if (bulkOps.length > 0) {
            await Product.bulkWrite(bulkOps);
            return res.json({ message: `成功更新 ${bulkOps.length} 筆調入資料` });
        } else {
            return res.json({ message: "無調入資料需要更新" });
        }

    } catch (err) {
        console.error(err);
        return res.status(500).json({ message: "抓取或更新調入資料失敗", error: err.message });
    }
});

// 抓取調出資料並更新資料庫
app.post('/api/fetchCallOutData/:storeName', async (req, res) => {
    const storeName = req.params.storeName;

    try {
        // 後端計算特殊系統年月
        const lastDay = new Date(year, formattedMonth, 0).getDate();

        const startDate = `${year}-${String(formattedMonth).padStart(2, '0')}-01`;
        const endDate = `${year}-${String(formattedMonth).padStart(2, '0')}-${lastDay}`;

        // 抓取 Kingzaap 調出資料
        const apiResponse = await axios.post(
            "https://kingzaap.unium.com.tw/BohAPI/MSCTTOMI/FindCallUpData",
            {
                Str_No: storeName, // 調出使用 Str_No
                StartTime: startDate,
                EndTime: endDate
            },
            {
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/plain, */*"
                }
            }
        );

        const data = apiResponse.data?.data || [];
        if (!Array.isArray(data)) return res.status(500).json({ message: "API 回傳格式錯誤" });

        // 生成品號 -> 調出數量映射
        const callOutMap = {};
        data.forEach(item => {
            const productCode = item.Goo_No?.trim();
            if (!productCode) return;

            // 累加同一品號的數量
            if (callOutMap[productCode]) {
                callOutMap[productCode] += Number(item.Qty || 0);
            } else {
                callOutMap[productCode] = Number(item.Qty || 0);
            }
        });

        // 更新 MongoDB
        const collectionName = `${year}${String(month).padStart(2, '0')}${storeName}`;
        const Product = mongoose.model(collectionName, productSchema);

        const bulkOps = [];
        for (const [productCode, qty] of Object.entries(callOutMap)) {
            const roundedFinalQty = parseFloat(qty.toFixed(2));

            bulkOps.push({
                updateOne: {
                    filter: { 品號: productCode },
                    update: { $set: { 調出: roundedFinalQty } }
                }
            });
        }

        if (bulkOps.length > 0) {
            await Product.bulkWrite(bulkOps);
            return res.json({ message: `成功更新 ${bulkOps.length} 筆調出資料` });
        } else {
            return res.json({ message: "無調出資料需要更新" });
        }

    } catch (err) {
        console.error(err);
        return res.status(500).json({ message: "抓取或更新調出資料失敗", error: err.message });
    }
});


// 新的 API 端點，處理上傳的進銷存 Excel 檔案
app.post('/api/uploadInventory/:storeName', upload.single('inventoryFile'), async (req, res) => {
    console.log('接收的請求:', req.body);
    console.log('請求文件:', req.file);

    const storeName = req.params.storeName;
    if (!req.file) {
        return res.status(400).json({ message: '請上傳 Excel 檔案' });
    }

    const uploadedFileName = req.file.originalname;
    console.log('上傳的文件名:', uploadedFileName);

    try {
        // 載入 Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);

        // 尋找工作表
        const worksheet = workbook.getWorksheet('總表');
        if (!worksheet) {
            return res.status(400).json({ message: '工作表「總表」不存在' });
        }

        // -------------------------------
        // 1. 自動尋找標題列
        // -------------------------------
        // 假設標題列在第 1 或第 2 行，先掃前 5 行找標題
        let headerRowNumber = null;
        let headerMap = {};  // {欄位名稱: index}

        for (let r = 1; r <= 5; r++) {
            const row = worksheet.getRow(r);
            const values = row.values.map(v => (v ? String(v).trim() : ''));

            // 判斷是否包含必要欄位，例如必須要有「品號」
            if (values.some(v => v.includes('品號'))) {
                headerRowNumber = r;
                values.forEach((val, idx) => {
                    if (!val) return;
                    headerMap[val] = idx; // 記錄標題與所在欄位
                });
                break;
            }
        }

        if (!headerRowNumber) {
            return res.status(400).json({ message: '未找到標題列（例如包含「品號」的列）' });
        }

        console.log('偵測到標題列在第', headerRowNumber, '行');
        console.log('欄位對應表:', headerMap);

        // -------------------------------
        // 2. 擷取資料列（標題列之後）
        // -------------------------------
        const dataRows = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= headerRowNumber) return; // 跳過標題之前的行

            const rowData = {};

            // 讀取所需欄位
            const getCellText = (c) => {
                const cellVal = row.getCell(c).value;
                if (cellVal && typeof cellVal === 'object' && cellVal.richText) {
                    return cellVal.richText.map(t => t.text).join('');
                }
                return cellVal !== null && cellVal !== undefined ? String(cellVal).trim() : '';
            };

            rowData.品號 = getCellText(headerMap['品號']);
            if (!rowData.品號) return; // 若品號空白則跳過該行

            rowData.廠商 = getCellText(headerMap['廠商']) || '未知';
            rowData.規格 = getCellText(headerMap['規格']) || '未知';
            rowData.盤點單位 = getCellText(headerMap['盤點單位']) || '未知';
            rowData.本月報價 = parseFloat(getCellText(headerMap['本月報價'])) || 0;
            rowData.進貨單位 = getCellText(headerMap['進貨單位']) || '未知';

            dataRows.push(rowData);
        });

        console.log(`共解析出 ${dataRows.length} 筆資料`);

        // -------------------------------
        // 3. 資料庫更新
        // -------------------------------
        const collectionName = `${year}${formattedMonth}${storeName}`;
        const Product = mongoose.models[collectionName] || mongoose.model(collectionName, productSchema);

        const bulkOps = [];

        // 先將本月報價全部重置為 0
        bulkOps.push({
            updateMany: {
                filter: {},
                update: { $set: { 本月報價: 0 } }
            }
        });

        // 將每一行資料加入批次更新
        dataRows.forEach(item => {
            bulkOps.push({
                updateOne: {
                    filter: { 品號: item.品號 },
                    update: {
                        $set: {
                            廠商: item.廠商,
                            規格: item.規格,
                            盤點單位: item.盤點單位,
                            本月報價: item.本月報價,
                            進貨單位: item.進貨單位
                        }
                    },
                    upsert: false
                }
            });
        });

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
        console.log('未收到上傳檔案');
        return res.status(400).json({ message: '請上傳本月進貨量文件' });
    }

    try {
        console.log('開始解析Excel檔案...');
        // 讀取 Excel
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        console.log(`找到工作表：${sheetName}，資料行數：${jsonData.length}`);

        if (!jsonData || jsonData.length < 2) {
            console.log('資料行數不足，檔案可能格式不正確');
            return res.status(400).json({ message: 'Excel 檔案格式不正確，缺少資料' });
        }

        // 清空本月進貨
        await Product.updateMany({}, { $set: { 本月進貨: 0, 進貨上傳: false } });

        const bulkOps = [];

        // 第一行是標題，從第2行開始
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];

            const productCode = (row[1] || '').toString().trim(); // 商品代號在第2欄
            const totalQty = parseFloat(row[10]) || 0;            // 總數量在第11欄

            if (!productCode) continue;

            console.log(`第 ${i + 1} 行：商品代號=${productCode}, 總數量=${totalQty}`);

            bulkOps.push({
                updateOne: {
                    filter: { 品號: productCode },
                    update: { $set: { 本月進貨: totalQty, 進貨上傳: true } }
                }
            });
        }

        if (bulkOps.length > 0) {
            const result = await Product.bulkWrite(bulkOps);
            console.log('批次更新完成：', result);

            // 廣播更新
            io.emit('monthlyPurchaseUpdated', { storeName });

            return res.status(200).json({
                message: `成功更新 ${bulkOps.length} 筆資料`,
                updated: bulkOps.length
            });
        } else {
            console.log('沒有資料需要更新');
            return res.status(200).json({ message: '無資料需要更新' });
        }
    } catch (error) {
        console.error('處理本月進貨量 Excel 檔案時發生錯誤:', error);
        res.status(500).json({ message: '處理檔案時發生錯誤', error: error.message });
    }
});

// 新的 API 端點，查詢本月使用量為負值的品項
app.get('/api/negativeUsageItems/:storeName', async (req, res) => {
    const storeName = req.params.storeName;
    const collectionName = `${year}${formattedMonth}${storeName}`;
    const Product = mongoose.model(collectionName, productSchema);

    try {
        const items = await Product.find({ 停用: false });

        // 計算使用量並過濾負值
        const result = items
            .map(item => {
                const 本月使用量 = (item.本月進貨 || 0)
                    + (item.期初盤點 || 0)
                    + (item.調入 || 0)
                    - (item.期末盤點 || 0)
                    - (item.調出 || 0);

                return {
                    品號: item.品號,
                    品名: item.品名,
                    本月進貨: item.本月進貨 || 0,
                    期初盤點: item.期初盤點 || 0,
                    期末盤點: item.期末盤點 || 0,
                    調入: item.調入 || 0,
                    調出: item.調出 || 0,
                    本月使用量
                };
            })
            .filter(item => item.本月使用量 < 0);

        res.json(result);
    } catch (err) {
        console.error(err);
        res.status(500).json({ message: '查詢失敗', error: err.message });
    }
});

// API端點: 檢查伺服器內部狀況
app.get('/api/checkConnections', (req, res) => {
    // 檢查服務器內部狀況，假設這裡始終有效
    res.status(200).json({ serverConnected: true });
});

// API 端點: 檢查 kingzaap 伺服器狀態
app.get('/api/ping', async (req, res) => {
    const url = 'https://kingzaap.unium.com.tw/BohWeb/';
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 5000); // 設定 5 秒超時

    try {
        const response = await fetch(url, {
            method: 'GET',
            signal: controller.signal // 使用 AbortController 處理超時
        });

        clearTimeout(timeoutId);

        // 如果 HTTP 狀態碼在 200-299 之間，則視為成功
        if (response.ok) {
            console.log('Kingzaap 伺服器連線成功');
            res.status(200).json({ KingzaapConnected: true });
        } else {
            console.error(`Kingzaap 伺服器回應錯誤狀態碼: ${response.status}`);
            res.status(response.status).json({
                KingzaapConnected: false,
                message: `伺服器回應狀態碼: ${response.status}`
            });
        }
    } catch (error) {
        clearTimeout(timeoutId);

        let errorMessage = '連線失敗';
        if (error.name === 'AbortError') {
            errorMessage = '連線超時，無法連線至 Kingzaap 伺服器';
        } else if (error.code === 'ENOTFOUND') {
            errorMessage = 'DNS 解析錯誤，無法找到 Kingzaap 伺服器';
        } else {
            errorMessage = `連線錯誤: ${error.message}`;
        }
        console.error(errorMessage);
        res.status(500).json({
            KingzaapConnected: false,
            message: errorMessage
        });
    }
});


// 更新產品數量的 API 端點
app.put('/api/products/:storeName/:productCode/quantity', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const { 期末盤點 } = req.body;
        const storeRoom = req.params.storeName;

        // 更新指定產品的數量
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 期末盤點: 期末盤點, 盤點完成: true, 最後更新時間: new Date(), 最後更新項目: '期末盤點' } }, // 更新內容
            { new: true }
        );

        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }
        // 檢查進貨上傳是否為 true
        if (updatedProduct.進貨上傳 === true) {
            // 計算本月用量
            const monthlyUsage =
                updatedProduct.本月進貨 +
                updatedProduct.期初盤點 +
                updatedProduct.調入 -
                updatedProduct.調出 -
                updatedProduct.期末盤點;

            // 如果本月用量為負數，廣播警告訊息給前端
            const productName = updatedProduct.品名;

            if (monthlyUsage < 0) {
                // 從更新後的產品中取得商品名稱
                // 建立包含商品名稱的警告訊息
                const alertMessage = `產品 ${productName} 的本月用量為負數，請檢查！`;

                // 廣播訊息，包含商品名稱
                io.to(storeName).emit('negativeUsageAlert', {
                    type: error,
                    message: alertMessage,
                    productName: productName // 將 productCode 替換為 productName
                });
            } else if (monthlyUsage > 300) {
                const alertMessage = `產品 ${productName} 的本月用量過高，請檢查月末盤點量！`;

                // 廣播訊息，包含商品名稱
                io.to(storeName).emit('negativeUsageAlert', {
                    type: info,
                    message: alertMessage,
                    productName: productName // 將 productCode 替換為 productName
                });
            }
        }
        // 廣播更新訊息給所有用戶
        io.to(storeName).emit('productUpdated', updatedProduct, storeRoom);

        res.json(updatedProduct);
    } catch (error) {
        console.error('更新產品時出錯:', error);
        res.status(400).send('更新失敗');
    }
});

// 標記為未盤點的 API 端點
app.post('/api/products/:storeName/:productCode/notInventoried', limiter, async (req, res) => {
    const storeName = req.params.storeName || 'notStart'; // 取得 URL 中的 storeName
    const collectionName = `${year}${formattedMonth}${storeName}`; // 根據年份、月份和門市產生集合品名
    const Product = mongoose.model(collectionName, productSchema);

    // 檢查商店品名是否有效
    if (storeName === 'notStart') {
        return res.status(400).send('門市錯誤'); // 使用 400 Bad Request 回傳錯誤
    }

    try {
        const { productCode } = req.params;
        const storeRoom = req.params.storeName;
        // 更新指定產品的數量
        const updatedProduct = await Product.findOneAndUpdate(
            { 品號: productCode },
            { $set: { 期末盤點: 0, 盤點完成: false, 最後更新時間: new Date(), 最後更新項目: '未盤點' } }, // 更新內容
            { new: true }
        );


        if (!updatedProduct) {
            return res.status(404).send('產品未找到');
        }


        // 成功後透過 Socket.IO 廣播給該門市所有連線
        if (global.io) {
            io.to(storeName).emit('updateInventoryStatus', {
                productIds,
                status: false
            });
        }

        return res.status(200).json({
            message: '盤點狀態已更新為未盤點',
            modifiedCount: updateResult.modifiedCount
        });
    } catch (error) {
        console.error('更新盤點狀態失敗:', error);
        return res.status(500).json({ message: '更新盤點狀態失敗', error: error.message });
    }
});

// 更新產品停用狀態的 API 端點
app.put('/api/products/:storeName/:productCode/depot', async (req, res) => {
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
app.put('/api/products/:storeName/:productCode/expiryDate', async (req, res) => {
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

// 批次更新產品庫別、廠商、停用狀態的 API 端點
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

// 公告 API 端點
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

// 處理 404 錯誤的中介軟體
app.use((req, res, next) => {
    // 設置 HTTP 狀態碼為 404
    res.status(404);

    // 建議：回傳一個最標準、最通用的 JSON 錯誤，不包含任何路由細節
    if (req.accepts('json')) {
        res.json({
            // 保持訊息簡潔，只說明找不到
            error: "Not Found",
            code: 404
        });
        return;
    }

    // 確保瀏覽器直接訪問時，也只返回一個通用的訊息
    res.type('text/plain').send('404 Not Found');
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
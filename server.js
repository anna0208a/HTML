/* // server.js
import express from 'express';
import multer from 'multer';
import cors from 'cors';
import { uploadPDF, generateExcel } from './pcrver.js';

const app = express();
const port = 3000;
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());

app.post('/api/analyze', upload.single('pdf'), async (req, res) => {
  try {
    const filepath = req.file.path;
    const result = await uploadPDF(filepath);
    res.json({ success: true, data: result });
  } catch (error) {
    console.error('❌ 分析錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/export-excel', async (req, res) => {
  try {
    const data = req.body;
    await generateExcel(data);
    res.json({ success: true });
  } catch (error) {
    console.error('❌ 匯出錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
}); */
// server.js
import express from 'express';
import multer from 'multer';
import cors from 'cors';
import { uploadPDF, generateExcel } from './pcrver.js';

const app = express();
const port = 3000;
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());

app.post('/api/analyze', upload.fields([
  { name: 'pcrFile', maxCount: 1 },
  { name: 'productFile', maxCount: 1 }
]), async (req, res) => {
  try {
    const pcrPath = req.files?.pcrFile?.[0]?.path;
    const productPath = req.files?.productFile?.[0]?.path;

    if (!pcrPath || !productPath) {
      throw new Error('缺少 pcr 或產品 PDF 檔案');
    }

    const result = await uploadPDF(pcrPath, productPath);
    res.json({ success: true, data: result });
  } catch (error) {
    console.error('❌ 分析錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/export-excel', async (req, res) => {
  try {
    const data = req.body;
    await generateExcel(data);
    res.json({ success: true });
  } catch (error) {
    console.error('❌ 匯出錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});

const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const { importToGoogleSheet } = require('./utils/excelParser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 25 * 1024 * 1024,
  },
  fileFilter: (req, file, cb) => {
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'text/csv'
    ];
    
    const allowedExts = /\.(xlsx|xls|csv)$/i;
    
    if (allowedMimes.includes(file.mimetype) || allowedExts.test(file.originalname)) {
      cb(null, true);
    } else {
      cb(new Error('Format file tidak didukung. Gunakan .xlsx, .xls, atau .csv'));
    }
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: 'Server running',
    timestamp: new Date().toISOString()
  });
});

app.post('/api/upload', upload.single('file'), async (req, res) => {

  req.setTimeout(900000);  
  res.setTimeout(900000);  
  
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'Tidak ada file yang diupload'
      });
    }

    console.log(`\n📊 Processing file: ${req.file.originalname}`);
    console.log(`   Size: ${(req.file.size / 1024).toFixed(2)} KB`);
    
    const result = await importToGoogleSheet(req.file);
    
    console.log(`✓ Import completed: ${result.rowsImported} rows\n`);
    
    let message = `Berhasil import ${result.rowsImported.toLocaleString()} baris data`;
    if (result.duplicatesSkipped > 0) {
      message += ` (${result.duplicatesSkipped.toLocaleString()} duplikat di-skip)`;
    }
    
    res.json({
      success: true,
      message: message,
      data: {
        rowsImported: result.rowsImported,
        duplicatesSkipped: result.duplicatesSkipped,
        duplicateIds: result.duplicateDetails ? result.duplicateDetails.map(d => d.idProyek) : [],
        startNumber: result.lastNumber - result.rowsImported + 1,
        endNumber: result.lastNumber,
        totalTime: result.totalTime 
      }
    });

  } catch (error) {
    console.error('❌ Upload error:', error.message);
    res.status(500).json({
      success: false,
      message: error.message || 'Terjadi kesalahan saat import data'
    });
  }
});

app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        success: false,
        message: 'File terlalu besar. Maksimal 10MB'
      });
    }
  }
  
  console.error('Server error:', error);
  res.status(500).json({
    success: false,
    message: error.message || 'Terjadi kesalahan pada server'
  });
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    message: 'Endpoint tidak ditemukan'
  });
});

app.listen(PORT, () => {
  console.log('\n========================================');
  console.log('🚀 Excel Import System');
  console.log('========================================');
  console.log(`📍 Server: http://localhost:${PORT}`);
  console.log(`📊 Status: Running`);
  console.log(`⏰ Started: ${new Date().toLocaleString('id-ID')}`);
  console.log('========================================\n');
});

module.exports = app;
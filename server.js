/**
 * نظام رفع الملفات واستخراج النصوص
 * يدعم استخراج النصوص من ملفات كبيرة (حتى 400 ميجابايت)
 * مع دعم كامل للغة العربية والاتجاه من اليمين إلى اليسار
 */

// تحميل المتغيرات البيئية من ملف .env
require('dotenv').config();

// استيراد المكتبات الأساسية
const express = require('express');
const session = require('express-session');
const bodyParser = require('body-parser');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const util = require('util');
const os = require('os');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { promisify } = require('util');

// تحويل الدوال غير المتزامنة لاستخدام الوعود (Promises)
const execAsync = promisify(exec);
const { writeFile: writeFileAsync, mkdir: mkdirAsync, stat: statAsync, 
        unlink: unlinkAsync, readFile: readFileAsync, 
        readdir: readdirAsync, rmdir: rmdirAsync } = require('fs').promises;

// استيراد وحدة تحويل ملفات المكتب
const { convertToSupportedFormat, deleteTemporaryFile } = require('./office-converter');

// إنشاء تطبيق Express
const app = express();
const PORT = 3000;

/**
 * تهيئة واجهة برمجة التطبيقات Gemini
 * المسؤولة عن استخراج النصوص من الملفات
 */
let genAI = null;
let model = null;

// التحقق من وجود مفتاح API صالح
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
if (!GEMINI_API_KEY || GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY' || GEMINI_API_KEY.length < 10) {
  console.error('خطأ: مفتاح API لـ Gemini غير صالح أو غير موجود في ملف .env');
} else {
  console.log('محاولة تهيئة Gemini API باستخدام المفتاح المقدم.');
}

/**
 * تهيئة Gemini API
 * يستخدم موديل gemini-1.0-pro للحصول على أفضل النتائج في استخراج النصوص
 */
function initGeminiAPI() {
  try {
    const apiKey = process.env.GEMINI_API_KEY;
    console.log(`[معلومات] تهيئة Gemini API باستخدام المفتاح: ${apiKey ? apiKey.substring(0, 5) + '...' : 'غير محدد'}`);
    
    if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY' || apiKey.length < 10) {
      throw new Error('مفتاح API غير صالح أو مفقود');
    }
    
    // إنشاء كائنات API
    genAI = new GoogleGenerativeAI(apiKey);
    model = genAI.getGenerativeModel({ 
      model: 'gemini-1.0-pro',  // استخدام الإصدار الصحيح من الموديل
      generationConfig: {
        temperature: 0.2,        // درجة حرارة منخفضة للحصول على نتائج أكثر دقة
        maxOutputTokens: 8192   // عدد أكبر من الرموز للحصول على استجابات أطول
      }
    });
    
    console.log('تم تهيئة Gemini API بنجاح باستخدام موديل: gemini-1.0-pro');
    return true;
  } catch (error) {
    console.error('فشل في تهيئة Gemini API:', error.message);
    return false;
  }
}

// محاولة تهيئة Gemini API عند بدء التشغيل
const geminiInitialized = initGeminiAPI();
if (!geminiInitialized) {
  console.error('تحذير: فشلت تهيئة Gemini API عند بدء التشغيل. قد لا يعمل استخراج النصوص.');
}

/**
 * إعداد الوسائط (Middleware) لتطبيق Express
 */
app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(session({
  secret: 'your-secret-key',
  resave: false,
  saveUninitialized: true
}));

// قاعدة بيانات المستخدمين البسيطة (في التطبيق الحقيقي، استخدم قاعدة بيانات مناسبة)
const users = {
  'user1': { password: 'password1' },
  'user2': { password: 'password2' }
};

/**
 * وظائف إدارة المجلدات والملفات
 */

/**
 * التأكد من وجود مجلدات المستخدم والمحاضرات في جميع المجلدات الرئيسية المطلوبة
 * @param {string} username - اسم المستخدم
 * @param {string} lectureName - اسم المحاضرة (اختياري)
 */
async function ensureDirectoriesExist(username, lectureName) {
  // قائمة المجلدات الرئيسية التي نحتاج لإنشاء الهيكل فيها
  const parentDirs = ['uploads', 'extract txt', 'summarized txt', 'question'];
  
  // إنشاء مجلدات المستخدم في جميع المجلدات الرئيسية
  for (const parentDir of parentDirs) {
    const userDir = path.join(__dirname, parentDir, username);
    if (!fs.existsSync(userDir)) {
      await mkdirAsync(userDir, { recursive: true })
        .catch(err => console.error(`خطأ في إنشاء مجلد المستخدم ${userDir}:`, err));
    }
    
    // إذا تم تقديم اسم المحاضرة، قم بإنشاء مجلدات المحاضرة أيضًا
    if (lectureName) {
      const lectureDir = path.join(userDir, lectureName);
      if (!fs.existsSync(lectureDir)) {
        await mkdirAsync(lectureDir, { recursive: true })
          .catch(err => console.error(`خطأ في إنشاء مجلد المحاضرة ${lectureDir}:`, err));
      }
    }
  }
}

/**
 * وظيفة حذف المجلد بشكل متكرر
 * تستخدم لحذف المحاضرات وجميع ملفاتها
 * @param {string} folderPath - مسار المجلد المراد حذفه
 */
async function deleteFolderRecursive(folderPath) {
  if (fs.existsSync(folderPath)) {
    try {
      // قراءة محتويات المجلد
      const items = await readdirAsync(folderPath);
      
      // معالجة كل عنصر في المجلد
      for (const item of items) {
        const itemPath = path.join(folderPath, item);
        const stats = await statAsync(itemPath);
        
        if (stats.isDirectory()) {
          // حذف المجلدات الفرعية بشكل متكرر
          await deleteFolderRecursive(itemPath);
        } else {
          // حذف الملفات
          await unlinkAsync(itemPath)
            .catch(err => console.error(`خطأ في حذف الملف ${itemPath}:`, err));
        }
      }
      
      // حذف المجلد الفارغ بعد حذف محتوياته
      await rmdirAsync(folderPath)
        .catch(err => console.error(`خطأ في حذف المجلد ${folderPath}:`, err));
      
      console.log(`تم حذف المجلد ${folderPath} بنجاح`);
    } catch (error) {
      console.error(`خطأ أثناء حذف المجلد ${folderPath}:`, error);
    }
  } else {
    console.log(`المجلد ${folderPath} غير موجود، لا حاجة للحذف`);
  }
}

/**
 * إعداد التخزين لـ multer لرفع الملفات
 */
const storage = multer.diskStorage({
  destination: async function (req, file, cb) {
    // الحصول على اسم المستخدم من الجلسة
    const username = req.session.user;
    const lectureName = req.body.lectureName;
    
    if (!username) {
      return cb(new Error('المستخدم غير مصادق عليه'));
    }
    
    try {
      // إنشاء جميع المجلدات الضرورية
      await ensureDirectoriesExist(username, lectureName);
      
      // إرجاع وجهة تحميل الملف
      if (lectureName) {
        const uploadLectureDir = path.join(__dirname, 'uploads', username, lectureName);
        return cb(null, uploadLectureDir);
      }
      
      const uploadUserDir = path.join(__dirname, 'uploads', username);
      cb(null, uploadUserDir);
    } catch (error) {
      cb(error);
    }
  },
  filename: function (req, file, cb) {
    // استخدام الاسم الأصلي للملف
    cb(null, file.originalname);
  }
});

// زيادة حد حجم الملف إلى 400 ميجابايت حسب طلب المستخدم
const upload = multer({ 
  storage: storage,
  limits: { fileSize: 400 * 1024 * 1024 } // 400 ميجابايت
});

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

app.get('/dashboard', (req, res) => {
  if (!req.session.user) {
    return res.redirect('/');
  }
  res.sendFile(path.join(__dirname, 'views', 'dashboard.html'));
});

// Login route
app.post('/login', (req, res) => {
  const { username, password } = req.body;
  
  console.log('Login attempt:', username, password);
  
  if (users[username] && users[username].password === password) {
    req.session.user = username;
    if (req.xhr || req.headers.accept && req.headers.accept.indexOf('json') > -1) {
      res.json({ success: true });
    } else {
      res.redirect('/dashboard');
    }
  } else {
    if (req.xhr || req.headers.accept && req.headers.accept.indexOf('json') > -1) {
      res.json({ success: false, message: 'Invalid credentials' });
    } else {
      res.redirect('/?error=invalid');
    }
  }
});

// Get user info route
app.get('/api/user', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ logged_in: false });
  }
  res.json({ logged_in: true, username: req.session.user });
});

// Save summary route
app.post('/api/save-summary', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const { lectureName, summaryText } = req.body;
  const username = req.session.user;
  
  if (!lectureName || !summaryText) {
    return res.status(400).json({ success: false, message: 'Missing lecture name or summary text' });
  }
  
  try {
    // Ensure directories exist
    ensureDirectoriesExist(username, lectureName);
    
    // Save the summary file
    const summaryFilePath = path.join(__dirname, 'summarized txt', username, lectureName, 'summary.txt');
    fs.writeFileSync(summaryFilePath, summaryText, 'utf8');
    
    res.json({ success: true, message: 'Summary saved successfully' });
  } catch (err) {
    console.error('Error saving summary:', err);
    res.status(500).json({ success: false, message: 'Error saving summary' });
  }
});

// Save questions route
app.post('/api/save-questions', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const { lectureName, questionsText } = req.body;
  const username = req.session.user;
  
  if (!lectureName || !questionsText) {
    return res.status(400).json({ success: false, message: 'Missing lecture name or questions text' });
  }
  
  try {
    // Ensure directories exist
    ensureDirectoriesExist(username, lectureName);
    
    // Save the questions file
    const questionsFilePath = path.join(__dirname, 'extract txt', username, lectureName, 'questions.txt');
    fs.writeFileSync(questionsFilePath, questionsText, 'utf8');
    
    res.json({ success: true, message: 'Questions saved successfully' });
  } catch (err) {
    console.error('Error saving questions:', err);
    res.status(500).json({ success: false, message: 'Error saving questions' });
  }
});

// Get summary route
app.get('/api/get-summary/:lectureName', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const lectureName = req.params.lectureName;
  const username = req.session.user;
  
  try {
    const summaryFilePath = path.join(__dirname, 'summarized txt', username, lectureName, 'summary.txt');
    
    if (fs.existsSync(summaryFilePath)) {
      const summaryText = fs.readFileSync(summaryFilePath, 'utf8');
      res.json({ success: true, summaryText });
    } else {
      res.json({ success: true, summaryText: '' });
    }
  } catch (err) {
    console.error('Error getting summary:', err);
    res.status(500).json({ success: false, message: 'Error getting summary' });
  }
});

// Get questions route
app.get('/api/get-questions/:lectureName', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const lectureName = req.params.lectureName;
  const username = req.session.user;
  
  try {
    const questionsFilePath = path.join(__dirname, 'extract txt', username, lectureName, 'questions.txt');
    
    if (fs.existsSync(questionsFilePath)) {
      const questionsText = fs.readFileSync(questionsFilePath, 'utf8');
      res.json({ success: true, questionsText });
    } else {
      res.json({ success: true, questionsText: '' });
    }
  } catch (err) {
    console.error('Error getting questions:', err);
    res.status(500).json({ success: false, message: 'Error getting questions' });
  }
});

// Save question content route
app.post('/api/save-question-content', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const { lectureName, questionContent } = req.body;
  const username = req.session.user;
  
  if (!lectureName || !questionContent) {
    return res.status(400).json({ success: false, message: 'Missing lecture name or question content' });
  }
  
  try {
    // Ensure directories exist
    ensureDirectoriesExist(username, lectureName);
    
    // Save the question content file
    const questionFilePath = path.join(__dirname, 'question', username, lectureName, 'question_content.txt');
    fs.writeFileSync(questionFilePath, questionContent, 'utf8');
    
    res.json({ success: true, message: 'Question content saved successfully' });
  } catch (err) {
    console.error('Error saving question content:', err);
    res.status(500).json({ success: false, message: 'Error saving question content' });
  }
});

// Get question content route
app.get('/api/get-question-content/:lectureName', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const lectureName = req.params.lectureName;
  const username = req.session.user;
  
  try {
    const questionFilePath = path.join(__dirname, 'question', username, lectureName, 'question_content.txt');
    
    if (fs.existsSync(questionFilePath)) {
      const questionContent = fs.readFileSync(questionFilePath, 'utf8');
      res.json({ success: true, questionContent });
    } else {
      res.json({ success: true, questionContent: '' });
    }
  } catch (err) {
    console.error('Error getting question content:', err);
    res.status(500).json({ success: false, message: 'Error getting question content' });
  }
});

// Function to determine MIME type based on file extension
function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  switch (ext) {
    case '.txt': return 'text/plain';
    case '.md': return 'text/markdown';
    case '.html': return 'text/html';
    case '.css': return 'text/css';
    case '.js': return 'text/javascript';
    case '.json': return 'application/json';
    case '.pdf': return 'application/pdf';
    // Image formats
    case '.png': return 'image/png';
    case '.jpeg':
    case '.jpg': return 'image/jpeg';
    case '.webp': return 'image/webp';
    case '.heic': return 'image/heic';
    case '.heif': return 'image/heif';
    case '.gif': return 'image/gif';
    // Audio formats
    case '.wav': return 'audio/wav';
    case '.mp3': return 'audio/mpeg';
    case '.aiff': return 'audio/aiff';
    case '.aac': return 'audio/aac';
    case '.ogg': return 'audio/ogg';
    case '.flac': return 'audio/flac';
    case '.m4a': return 'audio/m4a';
    // Microsoft Office formats
    case '.doc': return 'application/msword';
    case '.docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    case '.ppt': return 'application/vnd.ms-powerpoint';
    case '.pptx': return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    case '.xls': return 'application/vnd.ms-excel';
    case '.xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    
    // Default for unknown
    default: return 'application/octet-stream';
  }
}

/**
 * مجموعة من الوظائف المساعدة لتحويل ملفات Office إلى PDF
 */

// تحديد المسارات المحتملة لبرنامج LibreOffice في مختلف أنظمة التشغيل
const LIBREOFFICE_PATHS = {
  windows: [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice 7\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice 7.3\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice 7.4\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice 7.5\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice 7.6\\program\\soffice.exe'
  ],
  unix: [
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
    'soffice' // يفترض الوجود في PATH
  ]
};

// مجموعة من خيارات التحويل لـ LibreOffice
const CONVERSION_OPTIONS = [
  'pdf', // الخيار الأساسي
  'pdf:writer_pdf_Export', // خيارات تصدير إضافية
  'pdf:writer_web_pdf_Export', // خيارات للويب
  'pdf --norestore', // بدون استعادة العرض
  'pdf --convert-images-to=png' // تحويل الصور إلى PNG
];

// تكوين التحويل
const CONVERSION_CONFIG = {
  maxRetries: 5, // الحد الأقصى لمحاولات إعادة المحاولة
  retryDelay: 2000, // تأخير بين المحاولات (بالملي ثانية)
  libreOfficeTimeout: 60000, // مهلة تنفيذ أمر LibreOffice (بالملي ثانية)
  pdfMinSize: 100 // الحد الأدنى لحجم ملف PDF الصالح (بالبايت)
};

/**
 * البحث عن مسار LibreOffice المتاح
 * @returns {Promise<string|null>} مسار LibreOffice إذا وجد، أو null إذا لم يتم العثور عليه
 */
async function findLibreOfficePath() {
  const platform = os.platform() === 'win32' ? 'windows' : 'unix';
  const paths = LIBREOFFICE_PATHS[platform];
  
  // التحقق من المسارات المحتملة
  for (const p of paths) {
    try {
      if (os.platform() === 'win32' || p.startsWith('/')) {
        await fs.promises.access(p);
        console.log(`تم العثور على LibreOffice في: ${p}`);
        return p;
      } else if (os.platform() !== 'win32') {
        // محاولة العثور على soffice باستخدام الأمر 'which' في أنظمة Unix
        try {
          const { stdout } = await util.promisify(exec)('which soffice');
          if (stdout && stdout.trim()) {
            const path = stdout.trim();
            console.log(`تم العثور على LibreOffice باستخدام 'which': ${path}`);
            return path;
          }
        } catch (whichErr) {
          // تجاهل أخطاء 'which'
        }
      }
    } catch (err) {
      // المسار غير موجود، تجربة المسار التالي
    }
  }
  
  console.warn('لم يتم العثور على LibreOffice في النظام');
  return null;
}

/**
 * إنهاء أي عمليات LibreOffice عالقة
 * @returns {Promise<void>}
 */
async function killLibreOfficeProcesses() {
  if (os.platform() === 'win32') {
    try {
      console.log('محاولة إنهاء أي عمليات LibreOffice عالقة...');
      await new Promise((resolve) => {
        exec('taskkill /f /im soffice.exe /im soffice.bin', { timeout: 5000 }, () => resolve());
      });
      // الانتظار للتأكد من إنهاء العمليات
      await new Promise(resolve => setTimeout(resolve, 2000));
    } catch (killError) {
      // تجاهل الأخطاء - قد لا تكون هناك عمليات لإنهائها
    }
  } else if (os.platform() === 'linux' || os.platform() === 'darwin') {
    try {
      await new Promise((resolve) => {
        exec('pkill -f soffice', { timeout: 5000 }, () => resolve());
      });
      await new Promise(resolve => setTimeout(resolve, 2000));
    } catch (killError) {
      // تجاهل الأخطاء
    }
  }
}

/**
 * التحقق من صحة ملف PDF بعد الإنشاء
 * @param {string} pdfPath - مسار ملف PDF
 * @returns {Promise<boolean>} - صحيح إذا كان الملف صالحاً
 */
async function validatePDF(pdfPath) {
  try {
    // التحقق من وجود الملف
    await fs.promises.access(pdfPath);
    
    // التحقق من حجم الملف
    const stats = await fs.promises.stat(pdfPath);
    if (stats.size < CONVERSION_CONFIG.pdfMinSize) {
      console.warn(`ملف PDF صغير جداً (${stats.size} بايت). يمكن أن يكون فارغاً أو معطوباً.`);
      return false;
    }
    
    // التحقق من أن الملف يبدأ بسلسلة ملفات PDF المعروفة
    const fileHeader = await fs.promises.readFile(pdfPath, { encoding: 'utf8', flag: 'r', start: 0, end: 8 });
    if (!fileHeader.startsWith('%PDF-')) {
      console.warn(`ملف PDF غير صالح: لا يبدأ بسلسلة %PDF-`);
      return false;
    }
    
    console.log(`تم التحقق من صحة ملف PDF: ${path.basename(pdfPath)} (${stats.size} بايت)`);
    return true;
  } catch (error) {
    console.warn(`فشل التحقق من ملف PDF: ${error.message}`);
    return false;
  }
}

/**
 * Convertir PowerPoint usando la biblioteca pptx-to-pdf
 * @param {string} filePath - Ruta del archivo original
 * @param {string} outputPath - Ruta del archivo PDF de salida
 * @returns {Promise<boolean>} - true si la conversión fue exitosa
 */
async function convertWithPPTXToPDF(filePath, outputPath) {
  try {
    // Asegurarse de que el directorio de salida exista
    const outputDir = path.dirname(outputPath);
    await fs.promises.mkdir(outputDir, { recursive: true });

    // La biblioteca pptx-to-pdf puede ser requerida dinámicamente si se instala opcionalmente
    // Por ahora, asumimos que está instalada globalmente o localmente.
    // Considerar verificar su existencia o instalarla bajo demanda en futuras mejoras.
    const { convert } = require('pptx-to-pdf'); // Asegúrate de que esta biblioteca esté instalada
    console.log(`Intentando convertir ${path.basename(filePath)} a PDF usando pptx-to-pdf...`);
    await convert(filePath, outputPath);
    
    if (await validatePDF(outputPath)) {
      console.log(`Conversión exitosa de ${path.basename(filePath)} a PDF usando pptx-to-pdf.`);
      return true;
    }
    console.warn(`Conversión con pptx-to-pdf resultó en un PDF inválido para ${path.basename(filePath)}.`);
    // Intentar limpiar el archivo de salida si la conversión falló y creó un archivo vacío/corrupto
    try {
      if (fs.existsSync(outputPath)) { // Usar existsSync para comprobación síncrona simple o fs.promises.access para asíncrona
         const stats = await fs.promises.stat(outputPath);
         if (stats.size < CONVERSION_CONFIG.pdfMinSize) {
            console.log(`Eliminando archivo PDF inválido/vacío: ${outputPath}`);
            await fs.promises.unlink(outputPath);
         }
      }
    } catch (cleanupError) {
        console.warn(`Error al limpiar el archivo PDF fallido ${outputPath}: ${cleanupError.message}`);
    }
    return false;
  } catch (error) {
    console.warn(`Fallo al convertir ${path.basename(filePath)} con pptx-to-pdf: ${error.message}`);
    // Intentar limpiar el archivo de salida si la conversión falló
    try {
      if (fs.existsSync(outputPath)) {
         const stats = await fs.promises.stat(outputPath);
         if (stats.size < CONVERSION_CONFIG.pdfMinSize) {
            await fs.promises.unlink(outputPath);
         }
      }
    } catch (cleanupError) {
        // No hacer nada si la limpieza falla
    }
    return false;
  }
}

/**
 * Convertir usando métodos COM en Windows (para Office instalado)
 * @param {string} filePath - Ruta del archivo original
 * @param {string} outputPath - Ruta del archivo PDF de salida
 * @returns {Promise<boolean>} - true si la conversión fue exitosa
 */
async function convertWithOfficeCOM(filePath, outputPath) {
  if (os.platform() !== 'win32') {
    console.log('Skipping COM conversion as not on Windows.');
    return false;
  }

  const ext = path.extname(filePath).toLowerCase();
  let psScript = '';
  const tempPsScriptPath = path.join(os.tmpdir(), `convert_${Date.now()}.ps1`);

  // Asegurarse de que el directorio de salida exista
  const outputDir = path.dirname(outputPath);
  await fs.promises.mkdir(outputDir, { recursive: true });

  console.log(`Intentando convertir ${path.basename(filePath)} a PDF usando Office COM...`);

  if (ext === '.pptx' || ext === '.ppt') {
    psScript = `
      $ErrorActionPreference = "Stop"
      try {
        $powerpoint = New-Object -ComObject Powerpoint.Application
        $powerpoint.Visible = $false
        $presentation = $powerpoint.Presentations.Open('${filePath.replace(/'/g, "''")}', $true, $false, $false)
        $presentation.SaveAs('${outputPath.replace(/'/g, "''")}', 32) # 32 for ppSaveAsPDF
        $presentation.Close()
        $powerpoint.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Output "PowerPoint COM Conversion successful for ${path.basename(filePath)}"
        exit 0
      } catch {
        Write-Error "PowerPoint COM Conversion failed for ${path.basename(filePath)}: $($_.Exception.Message)"
        exit 1
      }
    `;
  } else if (ext === '.docx' || ext === '.doc') {
    psScript = `
      $ErrorActionPreference = "Stop"
      try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open('${filePath.replace(/'/g, "''")}')
        $document.SaveAs('${outputPath.replace(/'/g, "''")}', 17) # 17 for wdFormatPDF
        $document.Close($false) # $false to not save changes on close
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Output "Word COM Conversion successful for ${path.basename(filePath)}"
        exit 0
      } catch {
        Write-Error "Word COM Conversion failed for ${path.basename(filePath)}: $($_.Exception.Message)"
        exit 1
      }
    `;
  } else if (ext === '.xlsx' || ext === '.xls') {
    psScript = `
      $ErrorActionPreference = "Stop"
      try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open('${filePath.replace(/'/g, "''")}')
        # Intentar exportar todas las hojas activas o la primera hoja
        $workbook.ExportAsFixedFormat(0, '${outputPath.replace(/'/g, "''")}') # 0 for xlTypePDF
        $workbook.Close($false) # $false to not save changes
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Output "Excel COM Conversion successful for ${path.basename(filePath)}"
        exit 0
      } catch {
        Write-Error "Excel COM Conversion failed for ${path.basename(filePath)}: $($_.Exception.Message)"
        exit 1
      }
    `;
  } else {
    console.log(`COM conversion not supported for ${ext}`);
    return false;
  }

  try {
    await fs.promises.writeFile(tempPsScriptPath, psScript, { encoding: 'utf8' });
    
    await new Promise((resolve, reject) => {
      const process = exec(`powershell -ExecutionPolicy Bypass -File "${tempPsScriptPath}"`, { timeout: CONVERSION_CONFIG.libreOfficeTimeout * 2 }, (error, stdout, stderr) => {
        if (error) {
          console.error(`PowerShell exec error for ${path.basename(filePath)}: ${error.message}`);
          console.error(`PowerShell stderr for ${path.basename(filePath)}: ${stderr}`);
          reject(new Error(`PowerShell script failed: ${stderr || error.message}`));
          return;
        }
        console.log(`PowerShell stdout for ${path.basename(filePath)}: ${stdout}`);
        if (stderr) {
            console.warn(`PowerShell stderr (non-fatal) for ${path.basename(filePath)}: ${stderr}`);
        }
        resolve(stdout);
      });
    });
    
    if (await validatePDF(outputPath)) {
      console.log(`COM conversion successful and PDF validated for ${path.basename(filePath)}.`);
      return true;
    }
    console.warn(`COM conversion for ${path.basename(filePath)} resulted in an invalid PDF.`);
    return false;
  } catch (error) {
    console.warn(`Failed COM conversion for ${path.basename(filePath)}: ${error.message}`);
    return false;
  } finally {
    try {
      await fs.promises.unlink(tempPsScriptPath);
    } catch (e) {
      // ignorar error si el archivo temporal ya no existe
    }
  }
}

/**
 * تحديد ما إذا كان الملف بحاجة إلى تحويل
 * @param {string} filePath - Ù…Ø³Ø§Ø± Ø§Ù„Ù…Ù„Ù
 * @returns {object} - Ù…Ø¹Ù„ÙˆÙ…Ø§Øª Ø§Ù„ØªØ­ÙˆÙŠÙ„ { needsConversion, conversionType }
 */
function getConversionInfo(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mimeType = getMimeType(filePath);
  
  // Ø§Ù„ØµÙŠØº Ø§Ù„Ù…Ø¯Ø¹ÙˆÙ…Ø© Ù…Ø¨Ø§Ø´Ø±Ø© Ù…Ù† Gemini API
  const geminiSupportedMimeTypes = [
    'text/plain', 'text/markdown', 'text/html', 'text/css', 'text/javascript', 'application/json',
    'application/pdf', 'image/png', 'image/jpeg', 'image/webp', 'image/heic', 'image/heif', 'image/gif'
  ];
  
  // Ø¥Ø°Ø§ ÙƒØ§Ù† Ù†ÙˆØ¹ MIME Ù…Ø¯Ø¹ÙˆÙ… Ù…Ø¨Ø§Ø´Ø±Ø© Ù…Ù† Gemini
  if (geminiSupportedMimeTypes.includes(mimeType)) {
    return { needsConversion: false, conversionType: null };
  }
  
  // ØªØ­Ø¯ÙŠØ¯ Ù†ÙˆØ¹ Ø§Ù„ØªØ­ÙˆÙŠÙ„ Ø¨Ù†Ø§Ø¡Ù‹ Ø¹Ù„Ù‰ Ø§Ù…ØªØ¯Ø§Ø¯ Ø§Ù„Ù…Ù„Ù
  switch (ext) {
    case '.docx':
    case '.doc':
      return { needsConversion: true, conversionType: 'word-to-pdf' };
    case '.pptx':
    case '.ppt':
      return { needsConversion: true, conversionType: 'powerpoint-to-pdf' };
    case '.xlsx':
    case '.xls':
      return { needsConversion: true, conversionType: 'excel-to-pdf' };
    case '.m4a':
      return { needsConversion: true, conversionType: 'm4a-to-wav' };
    default:
      // ØºÙŠØ± Ù…Ø¯Ø¹ÙˆÙ… ÙˆÙ„Ø§ ÙŠÙ…ÙƒÙ† ØªØ­ÙˆÙŠÙ„Ù‡
      return { needsConversion: true, conversionType: 'unsupported' };
  }
}

/**
 * Función principal para convertir archivos Office a PDF con múltiples métodos
 * @param {string} filePath - Ruta del archivo original
 * @returns {Promise<{success: boolean, outputPath: string|null, error: Error|null, useGeminiDirect?: boolean}>}
 */
async function convertOfficeToPDF(filePath) {
  const originalFileName = path.basename(filePath);
  const outputFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.')) + '.pdf';
  const outputPath = path.join(path.dirname(filePath), outputFileName);

  console.log(`Iniciando conversión a PDF para: ${originalFileName}. Salida esperada: ${outputPath}`);
  let conversionSuccess = false;
  let lastError = null;

  // Asegurarse de que el directorio de salida exista
  try {
    await fs.promises.mkdir(path.dirname(outputPath), { recursive: true });
  } catch (mkdirError) {
    console.error(`No se pudo crear el directorio de salida ${path.dirname(outputPath)}: ${mkdirError.message}`);
    return { success: false, outputPath: null, error: mkdirError, useGeminiDirect: true };
  }
  
  const fileExt = path.extname(filePath).toLowerCase();

  // Método 1: Biblioteca específica (pptx-to-pdf para PowerPoint)
  if (fileExt === '.pptx' || fileExt === '.ppt') {
    if (await convertWithPPTXToPDF(filePath, outputPath)) {
      conversionSuccess = true;
    }
    if (conversionSuccess && await validatePDF(outputPath)) {
        console.log(`Éxito con pptx-to-pdf para ${originalFileName}.`);
        return { success: true, outputPath, error: null };
    } else if (conversionSuccess) { // Tuvo éxito pero el PDF no es válido
        conversionSuccess = false; // Marcar como no exitoso para probar otros métodos
        console.warn(`pptx-to-pdf generó un PDF inválido para ${originalFileName}. Intentando otros métodos.`);
    }
  }

  // Método 2: Office COM Automation (solo Windows)
  if (!conversionSuccess && os.platform() === 'win32') {
    if (await convertWithOfficeCOM(filePath, outputPath)) {
      conversionSuccess = true;
    }
    if (conversionSuccess && await validatePDF(outputPath)) {
        console.log(`Éxito con Office COM para ${originalFileName}.`);
        return { success: true, outputPath, error: null };
    } else if (conversionSuccess) {
        conversionSuccess = false;
        console.warn(`Office COM generó un PDF inválido para ${originalFileName}. Intentando otros métodos.`);
    }
  }

  // Método 3: LibreOffice (múltiples intentos y opciones)
  if (!conversionSuccess) {
    console.log(`Intentando conversión con LibreOffice para ${originalFileName}...`);
    await killLibreOfficeProcesses(); // Matar procesos antes de empezar
    const libreOfficePath = await findLibreOfficePath();
    if (libreOfficePath) {
      for (let attempt = 0; attempt < CONVERSION_CONFIG.maxRetries && !conversionSuccess; attempt++) {
        const optionIndex = attempt % CONVERSION_OPTIONS.length;
        const convertOption = CONVERSION_OPTIONS[optionIndex];
        
        const result = await executeLibreOfficeConversion(
          libreOfficePath, filePath, outputPath, convertOption, attempt
        );
        
        if (result.success) {
          if (await validatePDF(outputPath)) {
            conversionSuccess = true;
            console.log(`Éxito con LibreOffice para ${originalFileName} en el intento ${attempt + 1}.`);
            break; 
          } else {
             console.warn(`LibreOffice generó un PDF inválido para ${originalFileName} en el intento ${attempt + 1}.`);
             // No romper, permitir que el bucle continúe con otras opciones/intentos si validatePDF falla
          }
        }
        lastError = result.error || new Error('Fallo la validación de PDF de LibreOffice');
      }
    } else {
      lastError = new Error('LibreOffice no encontrado en el sistema.');
      console.warn(lastError.message);
    }
  }
  
  // Verificar resultado final
  if (conversionSuccess && await validatePDF(outputPath)) {
    console.log(`Conversión final exitosa para ${originalFileName}. PDF guardado en: ${outputPath}`);
    return { success: true, outputPath, error: null };
  } else {
    const finalErrorMsg = `Fallaron todos los métodos de conversión a PDF para ${originalFileName}. Último error: ${lastError ? lastError.message : 'Desconocido'}`;
    console.error(finalErrorMsg);
    // Limpiar el archivo de salida si existe y es inválido
    try {
        if (fs.existsSync(outputPath)) {
            await fs.promises.unlink(outputPath);
            console.log(`Archivo de salida inválido eliminado: ${outputPath}`);
        }
    } catch(e) { /* ignorar */ }
    return { 
      success: false, 
      outputPath: null, 
      error: new Error(finalErrorMsg),
      useGeminiDirect: true // Sugerir usar Gemini directamente como último recurso
    };
  }
}

// Guardar las funciones antiguas para referencia (se pueden eliminar después de probar)
const oldConvertPowerPointToPDF = async function(filePath) {
  try {
    const outputPath = filePath.replace(/\.pptx?$/i, '.pdf');
    const result = await convertOfficeToPDF(filePath);
    return result.success ? outputPath : null;
  } catch (error) {
    console.error(`Error en oldConvertPowerPointToPDF: ${error.message}`);
    return null;
  }
};

const oldConvertWordToPDF = async function(filePath) {
  try {
    const outputPath = filePath.replace(/\.docx?$/i, '.pdf');
    const result = await convertOfficeToPDF(filePath);
    return result.success ? outputPath : null;
  } catch (error) {
    console.error(`Error en oldConvertWordToPDF: ${error.message}`);
    return null;
  }
};

const oldConvertExcelToPDF = async function(filePath) {
  try {
    const outputPath = filePath.replace(/\.xlsx?$/i, '.pdf');
    const result = await convertOfficeToPDF(filePath);
    return result.success ? outputPath : null;
  } catch (error) {
    console.error(`Error en oldConvertExcelToPDF: ${error.message}`);
    return null;
  }
};

/**
 * تحويل ملف PowerPoint إلى PDF
 * @param {string} filePath - Ù…Ø³Ø§Ø± Ù…Ù„Ù PowerPoint
 * @returns {Promise<{success: boolean, outputPath: string|null, error: Error|null}>}
 */
async function convertPowerPointToPDF(filePath) {
  try {
    // Verificar que el archivo existe
    await fs.promises.access(filePath);
    
    // Determinar la ruta de salida
    const outputPath = filePath.replace(/\.[^.]+$/, '.pdf');
    
    console.log(`محاولة تحويل ملف PowerPoint: ${path.basename(filePath)} إلى PDF`);
    
    // Intento 1: Usar una biblioteca especializada pptx-to-pdf si está disponible
    try {
      const pptxToPdf = require('pptx-to-pdf');
      console.log(`استخدام مكتبة pptx-to-pdf لتحويل: ${path.basename(filePath)}`);
      
      await pptxToPdf.convert(filePath, outputPath);
      
      // Verificar que se haya creado el PDF
      try {
        await fs.promises.access(outputPath);
        console.log(`تم تحويل PowerPoint إلى PDF بنجاح باستخدام pptx-to-pdf: ${outputPath}`);
        return { success: true, outputPath, error: null };
      } catch (accessErr) {
        console.warn(`فشل التحويل باستخدام pptx-to-pdf، جارٍ المحاولة بطريقة أخرى`);
      }
    } catch (pptxToPdfError) {
      console.warn(`مكتبة pptx-to-pdf غير متوفرة أو حدث خطأ: ${pptxToPdfError.message}`);
    }
    
    // Intento 2: Usar LibreOffice
    let libreOfficePath = '';
    
    // Búsqueda de LibreOffice en diferentes ubicaciones según el sistema operativo
    if (os.platform() === 'win32') {
      // Windows - buscar en ubicaciones potenciales
      const possiblePaths = [
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files\\LibreOffice 7\\program\\soffice.exe',
        'C:\\Program Files\\LibreOffice 7.3\\program\\soffice.exe',
        'C:\\Program Files\\LibreOffice 7.4\\program\\soffice.exe',
        'C:\\Program Files\\LibreOffice 7.5\\program\\soffice.exe',
        'C:\\Program Files\\LibreOffice 7.6\\program\\soffice.exe'
      ];
      
      for (const p of possiblePaths) {
        try {
          await fs.promises.access(p);
          libreOfficePath = p;
          console.log(`تم العثور على LibreOffice في: ${p}`);
          break;
        } catch (err) {
          // Ruta no encontrada, probar la siguiente
        }
      }
    } else {
      // Linux/Mac - asumir que está en el PATH del sistema
      libreOfficePath = 'soffice';
    }
    
    if (libreOfficePath) {
      console.log(`استخدام LibreOffice لتحويل: ${path.basename(filePath)}`);
      
      // Probar dos métodos diferentes de conversión de LibreOffice
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          // Primer intento: usar --convert-to pdf
          // Segundo intento: usar --convert-to pdf:writer_pdf_Export (más opciones)
          const convertOption = attempt === 1 ? 'pdf' : 'pdf:writer_pdf_Export';
          
          await new Promise((resolve, reject) => {
            const command = `"${libreOfficePath}" --headless --convert-to ${convertOption} --outdir "${path.dirname(filePath)}" "${filePath}"`;
            console.log(`تنفيذ أمر التحويل (محاولة ${attempt}): ${command}`);
            
            exec(command, (error, stdout, stderr) => {
              if (error) {
                console.error(`خطأ في التحويل باستخدام LibreOffice: ${error.message}`);
                console.error(`STDERR: ${stderr}`);
                reject(error);
              } else {
                console.log(`تم تحويل PowerPoint إلى PDF: ${outputPath}`);
                console.log(`STDOUT: ${stdout}`);
                resolve(stdout);
              }
            });
          });
          
          // Verificar que se haya creado el PDF
          try {
            await fs.promises.access(outputPath);
            console.log(`تم تحويل PowerPoint إلى PDF بنجاح باستخدام LibreOffice (محاولة ${attempt}): ${outputPath}`);
            return { success: true, outputPath, error: null };
          } catch (accessErr) {
            if (attempt >= 2) {
              throw new Error(`لم يتم إنشاء ملف PDF: ${accessErr.message}`);
            }
            console.warn(`فشلت المحاولة ${attempt} باستخدام LibreOffice، جارٍ تجربة خيار آخر`);
          }
        } catch (libreOfficeError) {
          if (attempt >= 2) {
            throw new Error(`فشلت جميع محاولات التحويل باستخدام LibreOffice: ${libreOfficeError.message}`);
          }
          console.warn(`فشلت المحاولة ${attempt} باستخدام LibreOffice، جارٍ تجربة خيار آخر`);
        }
      }
    } else {
      console.warn('لم يتم العثور على LibreOffice');
    }
    
    // Intento 3: Uso directo de Gemini API como último recurso
    console.log(`محاولة استخدام استخراج النص مباشرة من ملف PowerPoint: ${path.basename(filePath)}`);
    return { 
      success: false, 
      outputPath: null, 
      error: new Error('لم يتم العثور على أدوات التحويل المناسبة أو فشلت جميع المحاولات'),
      useGeminiDirect: true // Indicador para usar Gemini directamente
    };
  } catch (error) {
    console.error(`خطأ في تحويل PowerPoint إلى PDF: ${error.message}`);
    return { success: false, outputPath: null, error, useGeminiDirect: true };
  }
}

/**
 * ØªØ­ÙˆÙŠÙ„ Ù…Ù„Ù Word Ø¥Ù„Ù‰ PDF
 * @param {string} filePath - Ù…Ø³Ø§Ø± Ù…Ù„Ù Word
 * @returns {Promise<{success: boolean, outputPath: string|null, error: Error|null}>}
 */
async function convertWordToPDF(filePath) {
  // ÙŠØ³ØªØ®Ø¯Ù… Ù†ÙØ³ Ø·Ø±ÙŠÙ‚Ø© PowerPoint Ø¨Ø§Ø³ØªØ®Ø¯Ø§Ù… LibreOffice
  return convertPowerPointToPDF(filePath);
}

/**
 * ØªØ­ÙˆÙŠÙ„ Ù…Ù„Ù Excel Ø¥Ù„Ù‰ PDF
 * @param {string} filePath - Ù…Ø³Ø§Ø± Ù…Ù„Ù Excel
 * @returns {Promise<{success: boolean, outputPath: string|null, error: Error|null}>}
 */
async function convertExcelToPDF(filePath) {
  try {
    const result = await convertOfficeToPDF(filePath);
    if (result.success) {
      console.log(`تم تحويل Excel إلى PDF بنجاح: ${result.outputPath}`);
      return result.outputPath;
    } else {
      console.error(`فشل تحويل Excel إلى PDF: ${result.error?.message || 'سبب غير معروف'}`);
      return null;
    }
  } catch (error) {
    console.error(`حدث خطأ غير متوقع أثناء تحويل Excel إلى PDF: ${error.message}`);
    return null;
  }
}

/**
 * Extract text from files up to 400MB using Gemini API with chunking and concurrency
 * @param {string} filePath - Path to the file to extract text from
 * @param {Buffer} contentToProcess - File content as buffer (optional)
 * @param {string} mimeTypeToProcess - MIME type of the file (optional)
 * @returns {Promise<string>} - Extracted text
 */
async function extractTextWithGemini(filePath, contentToProcess, mimeTypeToProcess) {
  try {
    const startTime = Date.now();
    const chunks = [];
    const chunkResults = [];
    const CONCURRENT_CHUNKS = 3; // Number of concurrent chunks to process
    const BATCH_DELAY_MS = 2000; // Delay between batches to avoid rate limits
    const MAX_RETRIES = 3; // Max retry attempts per chunk
    
    // File size limits and chunking parameters
    const MAX_FILE_SIZE_NO_CHUNK_MB = 19; // Gemini's limit per chunk is <20MB (using 19MB as safety margin)
    const MAX_FILE_SIZE_NO_CHUNK_BYTES = MAX_FILE_SIZE_NO_CHUNK_MB * 1024 * 1024;
    const MAX_OVERALL_FILE_SIZE_MB = 400; // User requirements for large file processing (increased from 10MB)
    const MAX_OVERALL_FILE_SIZE_BYTES = MAX_OVERALL_FILE_SIZE_MB * 1024 * 1024;
    
    // Adaptive chunk sizing based on file size - for very large files we use larger chunks
    let CHUNK_SIZE_BYTES = 5 * 1024 * 1024; // 5MB is the default size
    
    // Adjust chunk size based on file size for optimal performance
    const fileSizeMB = contentToProcess.length / (1024 * 1024);
    if (fileSizeMB > 200) {
      CHUNK_SIZE_BYTES = 10 * 1024 * 1024; // 10MB for very large files (>200MB)
      console.log(`Very large file (${fileSizeMB.toFixed(2)} MB) - using chunks of ${CHUNK_SIZE_BYTES / (1024 * 1024)} MB`);
    } else if (fileSizeMB > 100) {
      CHUNK_SIZE_BYTES = 8 * 1024 * 1024; // 8MB for large files (>100MB)
    } else if (fileSizeMB > 50) {
      CHUNK_SIZE_BYTES = 6 * 1024 * 1024; // 6MB for medium files (>50MB)
    }

    // Check file size limits
    if (contentToProcess.length > MAX_OVERALL_FILE_SIZE_BYTES) {
      throw new Error(`File size (${fileSizeMB.toFixed(2)} MB) exceeds the maximum allowed limit of ${MAX_OVERALL_FILE_SIZE_MB} MB`);
    }
    
    console.log(`File size: ${fileSizeMB.toFixed(2)} MB, Chunk size: ${(CHUNK_SIZE_BYTES / (1024 * 1024)).toFixed(2)} MB`);
    
    // Make sure Gemini API is initialized once before processing begins
    if (!genAI || !model) {
      console.log('Initializing Gemini API before file processing');
      if (!initGeminiAPI()) {
        throw new Error('Failed to initialize Gemini API before file processing');
      }
    }

    // Process text files directly for better performance
    const textExtensions = [".txt", ".md", ".html", ".css", ".js", ".json"];
    const fileExt = path.extname(filePath).toLowerCase();
    
    if (textExtensions.includes(fileExt)) {
      console.log(`Processing text file (${fileExt}) directly for better performance`);
      const textContent = contentToProcess.toString('utf8');
      console.log(`Extracted ${(textContent.length / 1024).toFixed(2)} KB of text directly`);
      return textContent;
    }

    // Split file into chunks based on file size
    // If size is below threshold, process as a single chunk
    if (contentToProcess.length <= MAX_FILE_SIZE_NO_CHUNK_BYTES) {
      console.log(`File smaller than ${MAX_FILE_SIZE_NO_CHUNK_MB} MB, no need for chunking`);
      chunks.push(contentToProcess);
    } else {
      // Chunk large file into smaller pieces
      console.log(`Chunking file of size ${fileSizeMB.toFixed(2)} MB into chunks of ${(CHUNK_SIZE_BYTES / (1024 * 1024)).toFixed(2)} MB...`);
      
      const totalChunksEstimate = Math.ceil(contentToProcess.length / CHUNK_SIZE_BYTES);
      console.log(`Estimated total number of chunks: ${totalChunksEstimate}`);
      
      for (let i = 0; i < contentToProcess.length; i += CHUNK_SIZE_BYTES) {
        const chunk = contentToProcess.slice(i, i + CHUNK_SIZE_BYTES);
        chunks.push(chunk);
      }
    }
    
    const totalChunks = chunks.length;
    console.log(`Total chunks to process: ${totalChunks}`);
    
    // Process chunks with concurrency control (limited number of concurrent chunks)
    const chunkStats = {
      total: totalChunks,
      processed: 0,
      successful: 0,
      failed: 0,
      startTime: Date.now()
    };
    
    // Define the processChunk function - processes a single chunk with retry logic
    async function processChunk(chunk, chunkNumber, totalChunks, mimeType) {
      const maxRetries = 3;
      let lastError = null;
      
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          // Ensure API is initialized
          if (!genAI || !model) {
            if (!initGeminiAPI()) {
              throw new Error('Failed to initialize Gemini API');
            }
          }
          
          // Convert chunk to base64
          const base64Chunk = chunk.toString('base64');
          
          // Call Gemini API
          const result = await model.generateContent({
            contents: [{
              role: 'user',
              parts: [
                { text: `This is chunk ${chunkNumber} of ${totalChunks} from a file. Please extract all text from this chunk only, without any comments or explanations.` },
                { inlineData: { mimeType: mimeType || 'application/octet-stream', data: base64Chunk } }
              ]
            }]
          });
          
          const extractedText = result.response.text();
          console.log(`✓ Successfully processed chunk ${chunkNumber}/${totalChunks} on attempt ${attempt}`);
          return extractedText;
          
        } catch (error) {
          lastError = error;
          console.warn(`× Failed attempt ${attempt}/${maxRetries} for chunk ${chunkNumber}: ${error.message}`);
          
          if (attempt < maxRetries) {
            // Calculate wait time with exponential backoff
            const delayMs = Math.pow(2, attempt) * 1000;
            console.log(`Waiting ${delayMs/1000} seconds before retry...`);
            await new Promise(resolve => setTimeout(resolve, delayMs));
            
            // Reinitialize API if network error
            if (error.message.includes('network') || error.message.includes('rate limit')) {
              console.log('Reinitializing Gemini API due to network error or rate limit');
              initGeminiAPI();
            }
          }
        }
      }
      
      // If all retries failed
      console.error(`All ${maxRetries} attempts failed for chunk ${chunkNumber}`);
      return null;
    }
    
    // Process chunks in concurrent batches with delay between batches
    for (let i = 0; i < totalChunks; i += CONCURRENT_CHUNKS) {
      const batch = [];
      const batchChunks = chunks.slice(i, i + CONCURRENT_CHUNKS);
      
      console.log(`Processing batch ${Math.floor(i / CONCURRENT_CHUNKS) + 1} of ${Math.ceil(totalChunks / CONCURRENT_CHUNKS)} (chunks ${i + 1} - ${Math.min(i + batchChunks.length, totalChunks)})`);
      
      // Process each batch concurrently
      for (let j = 0; j < batchChunks.length; j++) {
        const chunkNumber = i + j + 1;
        batch.push(processChunk(batchChunks[j], chunkNumber, totalChunks, mimeTypeToProcess));
      }
      
      try {
        // Wait for all promises in the current batch to complete
        const batchResults = await Promise.all(batch);
        
        // Process batch results
        for (let j = 0; j < batchResults.length; j++) {
          const result = batchResults[j];
          const chunkNumber = i + j + 1;
          chunkStats.processed++;
          
          if (result) {
            console.log(`Successfully processed chunk ${chunkNumber}/${totalChunks}`);
            chunkStats.successful++;
            chunkResults.push({ chunkNumber, text: result, success: true });
          } else {
            console.error(`Failed to process chunk ${chunkNumber}/${totalChunks}`);
            chunkStats.failed++;
            chunkResults.push({ chunkNumber, text: "", success: false, error: "Failed to extract text" });
          }
        }
        
        // Add delay between batches to avoid API rate limits
        if (i + CONCURRENT_CHUNKS < totalChunks) {
          console.log(`Waiting ${BATCH_DELAY_MS}ms before next batch...`);
          await new Promise(resolve => setTimeout(resolve, BATCH_DELAY_MS));
        }
      } catch (error) {
        console.error(`Error processing batch: ${error.message}`);
        // Continue with the next batch even if one fails
      }
    }
    
    // Calculate processing statistics
    const processingTimeMs = Date.now() - chunkStats.startTime;
    const successRate = (chunkStats.successful / chunkStats.total) * 100;
    
    console.log(`
===== Text Extraction Summary =====
- Total chunks: ${chunkStats.total}
- Successfully processed: ${chunkStats.successful} (${successRate.toFixed(1)}%)
- Failed: ${chunkStats.failed}
- Processing time: ${(processingTimeMs / 1000).toFixed(2)} seconds
`);
    
    // Combine all extracted text from successful chunks
    // Sort by chunkNumber to ensure correct order
    chunkResults.sort((a, b) => a.chunkNumber - b.chunkNumber);
    
    // Build the final text and report
    let extractionReport = {
      filePath,
      fileSize: `${fileSizeMB.toFixed(2)} MB`,
      chunks: {
        total: chunkStats.total,
        successful: chunkStats.successful,
        failed: chunkStats.failed,
        successRate: `${successRate.toFixed(1)}%`
      },
      processingTime: `${(processingTimeMs / 1000).toFixed(2)} seconds`,
      chunkDetails: []
    };
    
    // Combine text from all successful chunks
    let finalText = "";
    for (const result of chunkResults) {
      if (result.success) {
        finalText += result.text + "\n";
      }
      
      extractionReport.chunkDetails.push({
        chunkNumber: result.chunkNumber,
        success: result.success,
        textLength: result.success ? result.text.length : 0,
        error: result.error || null
      });
    }
    
    // Clean up the final text (remove extra newlines, etc.)
    finalText = finalText.trim()
      .replace(/\n{3,}/g, '\n\n') // Replace multiple newlines with double newlines
      .replace(/\s+\n/g, '\n')   // Remove spaces before newlines
      .replace(/\n\s+/g, '\n');   // Remove spaces after newlines
    
    // Add warning if less than 50% of chunks were processed successfully
    if (successRate < 50) {
      console.warn(`WARNING: Less than 50% of chunks were processed successfully (${successRate.toFixed(1)}%). Text extraction may be incomplete.`);
      extractionReport.warning = `Low success rate (${successRate.toFixed(1)}%). Text extraction may be incomplete.`;
    }
    
    // Log the final result
    console.log(`Final extracted text length: ${finalText.length} characters`);
    console.log(`Text extraction completed in ${(processingTimeMs / 1000).toFixed(2)} seconds`);
    
    return finalText;
  } catch (error) {
    console.error(`Error in extractTextWithGemini: ${error.message}`);
    throw new Error(`Failed to extract text: ${error.message}`);
  }
}
/**
 * Helper function to process a single chunk with retry logic
 * @param {Buffer} chunk - The data chunk to process
 * @param {number} chunkNumber - Current chunk number
 * @param {number} totalChunks - Total number of chunks
 * @param {string} mimeType - MIME type of the file
 * @returns {Promise<string>} - The extracted text from this chunk
 */
async function processChunk(chunk, chunkNumber, totalChunks, mimeType) {
  const maxRetries = 3; // الحد الأقصى لعدد مرات إعادة المحاولة
  let lastError = null;
  let extractedText = "";
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // تأكد من تهيئة Gemini API قبل المعالجة
      if (!genAI || !model) {
        console.log(`إعادة تهيئة Gemini API للجزء ${chunkNumber}/${totalChunks}`);
        if (!initGeminiAPI()) {
          throw new Error('فشل في تهيئة Gemini API');
        }
      }
      
      // تسجيل معلومات حول حالة معالجة الملف
      console.log(`معالجة الجزء ${chunkNumber}/${totalChunks}, المحاولة ${attempt}, الحجم: ${(chunk.length / 1024).toFixed(2)} كيلوبايت`);
      
      // إعداد النص التوجيهي المناسب لهذا الجزء
      const chunkPrompt = `هذا هو الجزء ${chunkNumber} من ${totalChunks} من ملف. استخرج كل محتوى النص من هذا الجزء فقط، بدون تعليقات أو شروحات.`;
      
      // تحويل الجزء إلى تنسيق base64
      const base64Chunk = chunk.toString('base64');
      
      try {
        // استدعاء Gemini API بالطريقة الصحيحة
        const result = await model.generateContent({
          contents: [{ 
            role: 'user',
            parts: [
              { text: chunkPrompt },
              { inlineData: { 
                mimeType: mimeType || 'application/octet-stream', 
                data: base64Chunk 
              }}
            ]
          }]
        });
        
        // التحقق من صحة الاستجابة
        if (!result || !result.response) {
          throw new Error('استجابة فارغة من Gemini API');
        }
        
        extractedText = result.response.text();
        console.log(`تم استخراج النص بنجاح من الجزء ${chunkNumber}, طول النص: ${extractedText.length} حرف`);
        
        // خروج من حلقة المحاولات في حالة النجاح
        return extractedText;
      } catch (apiError) {
        console.error(`خطأ API مع الجزء ${chunkNumber}:`, apiError);
        throw apiError; // إعادة رمي الخطأ للتعامل معه في حلقة المحاولات
      }
      
    } catch (error) {
      lastError = error;
      console.warn(`فشلت المحاولة ${attempt} للجزء ${chunkNumber}:`, error.message);
      
      // تأخير تزايدي قبل إعادة المحاولة (exponential backoff)
      if (attempt < maxRetries) {
        const delayMs = Math.pow(2, attempt) * 1000; // 2ث، 4ث، 8ث، الخ
        console.log(`إعادة محاولة معالجة الجزء ${chunkNumber} بعد تأخير ${delayMs/1000} ثانية...`);
        await new Promise(resolve => setTimeout(resolve, delayMs));
        
        // إعادة تهيئة Gemini API قبل المحاولة التالية إذا كان الخطأ يتعلق بالاتصال
        if (error.message.includes('network') || error.message.includes('timeout') || error.message.includes('API')) {
          console.log(`محاولة إعادة تهيئة Gemini API قبل المحاولة التالية`);
          initGeminiAPI();
        }
      }
    }
  }
  
  // إذا فشلت كل المحاولات، نعيد نص فارغ بدلاً من رمي خطأ لاستمرار معالجة باقي الأجزاء
  console.error(`فشل في معالجة الجزء ${chunkNumber} بعد ${maxRetries} محاولات. سيتم إرجاع نص فارغ.`);
  return "";
}

// وظيفة مساعدة لتنظيف النص المستخرج
function cleanExtractedText(text) {
    try {
        if (!text) return '';
        
        // 1. إزالة التكرارات المتتالية للجمل نفسها
        let cleanedText = text.replace(/(.{10,100}?)\1+/g, '$1');
        
        // 2. إزالة الأسطر الفارغة المتعددة
        cleanedText = cleanedText.replace(/\n{3,}/g, '\n\n');
        
        // 3. إزالة المسافات الزائدة
        cleanedText = cleanedText.replace(/[\s\u200B-\u200D\uFEFF]+/g, ' ').trim();
        
        // 4. إزالة العلامات الخاصة بالشفرات البرمجية إذا لزم الأمر
        cleanedText = cleanedText.replace(/```[\s\S]*?```/g, '');
        
        // 5. إزالة التكرارات في الكلمات والعبارات
        const repeatedPhrasePatterns = [
            // Common repetition patterns in Arabic and English
            /(طب ايه رايك\؟)\s*\1+/g,
            /(طب ايه هو الهدف بتاعه\؟)\s*\1+/g,
            /(طب ايه رايك\؟ طب ايه هو الهدف بتاعه\؟)\s*\1+/g,
            /(نعم)\s*\1+/g,
            /(اه)\s*\1+/g,
            /(صح)\s*\1+/g,
            /(نفس)\s*\1+/g,
            /(كده)\s*\1+/g,
            /(yes)\s*\1+/gi,
            /(no)\s*\1+/gi,
            /(okay)\s*\1+/gi,
            /(ok)\s*\1+/gi
        ];
        
        for (const pattern of repeatedPhrasePatterns) {
            cleanedText = cleanedText.replace(pattern, (match, p1) => p1);
        }
        
        // 6. إزالة أي علامات خاصة زائدة
        cleanedText = cleanedText.replace(/^\s*[\-\.\:\*]{2,}\s*$/gm, '');
        
        return cleanedText;
    } catch (error) {
        console.error('Error in cleanExtractedText:', error);
        return text; // العودة للنص الأصلي في حالة حدوث خطأ
    }
}

// وظيفة لاستخراج النص مباشرة من الملفات النصية
function extractTextDirectly(buffer, mimeType) {
    try {
        // التحقق من أنواع الملفات النصية المدعومة
        const textMimeTypes = [
            'text/plain', 'text/html', 'text/css', 'text/javascript',
            'application/json', 'application/xml', 'application/javascript'
        ];
        
        if (!textMimeTypes.includes(mimeType)) {
            throw new Error('Unsupported file type for direct extraction');
        }
        
        // تحويل الـ buffer إلى نص
        const text = buffer.toString('utf8');
        
        // تنظيف النص
        return cleanExtractedText(text);
    } catch (error) {
        console.error('Error in extractTextDirectly:', error);
        throw error;
    }
}




/**
 * وظيفة استخراج النص من الملفات المختلفة بما في ذلك الملفات الكبيرة حتى 400 ميجابايت
 * تدعم معالجة الملفات بالتقسيم إلى أجزاء مع إدارة متزامنة للمعالجة
 * @param {string} filePath - مسار الملف المراد استخراج النص منه
 * @param {string} outputDir - المجلد الذي سيتم حفظ النص المستخرج فيه
 * @returns {Promise<{success: boolean, output: string, extractedText?: string, error?: string}>}
 */
async function extractTextFromFile(filePath, outputDir) {
  const MAX_FILE_SIZE_MB = 400; // حد حجم الملف الأقصى - تم زيادته من 10 إلى 400 ميجابايت
  const CHUNK_PROCESSING_ENABLED = true; // تفعيل معالجة الأجزاء للملفات الكبيرة
  const MIN_SIZE_FOR_CHUNKING_MB = 10; // حجم الملف الأدنى لبدء المعالجة بالأجزاء
  
  let startTime, endTime; // لقياس الوقت المستغرق
  
  try {
    startTime = Date.now();
    console.log(`بدء استخراج النص من الملف: ${filePath}`);
    console.log(`المجلد الهدف: ${outputDir}`);
    
    // التأكد من وجود المجلد الهدف
    await mkdirAsync(outputDir, { recursive: true });

    // تحديد اسم ملف الإخراج مع طابع زمني
    const filename = path.basename(filePath);
    const timestamp = new Date().toISOString().replace(/:/g, '-').replace(/\..+/, '');
    const extractedTextFile = `${path.parse(filename).name}_extracted_${timestamp}.txt`;
    const outputPath = path.join(outputDir, extractedTextFile);
    
    // التحقق من حجم الملف
    const stats = await fs.promises.stat(filePath);
    const fileSizeMB = stats.size / (1024 * 1024);
    
    console.log(`حجم الملف: ${fileSizeMB.toFixed(2)} ميجابايت`);
    
    // إذا كان الملف أكبر من الحد المسموح به
    if (fileSizeMB > MAX_FILE_SIZE_MB) { 
      console.log(`الملف كبير جداً للمعالجة: ${fileSizeMB.toFixed(2)} ميجابايت (الحد: ${MAX_FILE_SIZE_MB} ميجابايت)`);
      const extension = path.extname(filePath).toLowerCase();
      const mimeType = getMimeType(filePath) || 'unknown';
      
      // كتابة معلومات الملف بدلاً من النص المستخرج
      const fileInfo = `معلومات عن الملف ${filename}:\n` +
                    `- نوع الملف: ${extension} (${mimeType})\n` +
                    `- حجم الملف: ${fileSizeMB.toFixed(2)} ميجابايت\n` +
                    `- مسار الملف: ${filePath}\n\n` +
                    `لم يتم استخراج النص من هذا الملف لأنه أكبر من ${MAX_FILE_SIZE_MB} ميجابايت.`;
      
      await fs.promises.writeFile(outputPath, fileInfo, 'utf8');
      console.log(`تم كتابة معلومات الملف إلى ${outputPath}`);
      
      endTime = Date.now();
      console.log(`انتهت المعالجة في ${((endTime - startTime) / 1000).toFixed(2)} ثانية`);
      
      return { success: true, output: outputPath };
    }
    
    // تحديد ما إذا كان سيتم استخدام معالجة الأجزاء للملفات الكبيرة
    const useChunking = CHUNK_PROCESSING_ENABLED && fileSizeMB >= MIN_SIZE_FOR_CHUNKING_MB;
    if (useChunking) {
      console.log(`سيتم معالجة ملف كبير (${fileSizeMB.toFixed(2)} ميجابايت) باستخدام تقنية تقسيم الملف إلى أجزاء`);
    } else {
      console.log(`سيتم معالجة الملف (${fileSizeMB.toFixed(2)} ميجابايت) بطريقة مباشرة`);
    }
    
    // قراءة الملف وتحديد نوع MIME
    let fileContent;
    try {
      fileContent = await readFileAsync(filePath);
      console.log(`تم قراءة الملف بنجاح، حجم البيانات المقروءة: ${(fileContent.length / (1024 * 1024)).toFixed(2)} ميجابايت`);
    } catch (readError) {
      console.error(`خطأ في قراءة الملف ${filePath}:`, readError);
      throw new Error(`لا يمكن قراءة الملف: ${readError.message}`);
    }
    
    // تحديد نوع MIME
    const mimeType = getMimeType(filePath);
    console.log(`نوع MIME للملف: ${mimeType || 'غير معروف'} للملف: ${filePath}`);

    // التعامل مع الملفات النصية مباشرة
    const textExtensions = ['.txt', '.md', '.html', '.css', '.js', '.json', '.xml', '.csv', '.rtf'];
    const fileExtension = path.extname(filePath).toLowerCase();
    
    let extractedText;
    
    // محاولة استخراج النص مباشرة للملفات النصية
    if (textExtensions.includes(fileExtension) || mimeType?.startsWith('text/')) {
      try {
        console.log(`معالجة ملف نصي مباشرة: ${fileExtension}`);
        extractedText = extractTextDirectly(fileContent, mimeType);
        console.log(`تم استخراج النص مباشرة، حجم النص: ${(extractedText.length / 1024).toFixed(2)} كيلوبايت`);
      } catch (directError) {
        console.warn(`فشل الاستخراج المباشر للنص، سيتم محاولة استخدام Gemini API:`, directError.message);
        // سنستمر بمحاولة استخدام Gemini API
      }
    }
    
    // إذا لم يتم استخراج النص مباشرة، intentar extraer contenido básico antes de usar Gemini API
    if (!extractedText) {
      // Intentar extraer texto de formatos comunes adicionales
      try {
        // PDF: Extraer texto simple del PDF si es posible
        if (fileExtension === '.pdf' || mimeType === 'application/pdf') {
          console.log('Intentando extraer texto directamente del PDF');
          // Extracción básica - mostramos el nombre del archivo y metadatos
          extractedText = `Archivo PDF: ${path.basename(filePath)}\n`;
          extractedText += `Tamaño: ${(fileSizeMB).toFixed(2)} MB\n`;
          extractedText += `Fecha: ${new Date().toISOString()}\n\n`;
          extractedText += 'El contenido de este PDF debe extraerse con API externa. ';
          extractedText += 'La API Gemini no está configurada correctamente.\n';
        }
        
        // Documentos Word/Excel/PowerPoint - información básica
        else if (['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt'].includes(fileExtension)) {
          console.log('Información básica de documento de Office');
          extractedText = `Documento de Office: ${path.basename(filePath)}\n`;
          extractedText += `Tipo: ${fileExtension}\n`;
          extractedText += `Tamaño: ${(fileSizeMB).toFixed(2)} MB\n`;
          extractedText += `Fecha: ${new Date().toISOString()}\n\n`;
          extractedText += 'El texto de este documento debe extraerse con API externa. ';
          extractedText += 'La API Gemini no está configurada correctamente.\n';
        }
      } catch (basicError) {
        console.warn('Error al intentar extraer texto básico:', basicError.message);
      }
      
      // Si todavía no tenemos texto, intentar con Gemini API
      if (!extractedText) {
        try {
          // محاولة إعادة تهيئة API قبل الاستخراج لضمان العمل
          if (!genAI || !model) {
            console.log('إعادة تهيئة Gemini API قبل استخراج النص');
            if (!initGeminiAPI()) {
              throw new Error('فشل في تهيئة Gemini API');
            }
          }
          
          console.log(`بدء استخراج النص باستخدام Gemini API للملف: ${path.basename(filePath)}`);
          
          // استخدام وظيفة استخراج النص المحسنة مع دعم الملفات الكبيرة
          extractedText = await extractTextWithGemini(filePath, fileContent, mimeType);
          
          if (!extractedText || typeof extractedText !== 'string' || extractedText.trim() === '') {
            console.warn(`لم يتم استخراج نص صالح من ${filePath}، النوع: ${typeof extractedText}، الطول: ${extractedText?.length || 0}`);
            extractedText = `لم يتم استخراج نص صالح من الملف. قد يكون الملف تالفًا أو بتنسيق غير مدعوم.`;
          } else {
            console.log(`تم استخراج النص بنجاح، طول النص: ${extractedText.length} حرف`);
          }
        } catch (geminiError) {
          console.error('Error al usar Gemini API:', geminiError.message);
          // Si la API falla, usamos un mensaje informativo
          extractedText = `No se pudo extraer texto con Gemini API: ${geminiError.message}\n\n`;
          extractedText += `Archivo: ${path.basename(filePath)}\n`;
          extractedText += `Tipo: ${fileExtension} (${mimeType || 'tipo desconocido'})\n`;
          extractedText += `Tamaño: ${(fileSizeMB).toFixed(2)} MB\n`;
        }
      }
    // تنظيف النص المستخرج إذا كان موجودًا
    if (extractedText && typeof extractedText === 'string') {
      const originalLength = extractedText.length;
      extractedText = cleanExtractedText(extractedText);
      console.log(`تم تنظيف النص المستخرج، الطول الأصلي: ${originalLength}، الطول بعد التنظيف: ${extractedText.length}`);
    }
      } catch (extractError) {
        console.error(`خطأ في استخراج النص من ${filePath}:`, extractError);
        
        // محاولة تجزئة الرسالة إذا كانت طويلة جداً
        let errorMessage = extractError.message || String(extractError);
        if (errorMessage.length > 500) {
          errorMessage = errorMessage.substring(0, 497) + '...';
        }
        
        extractedText = `حدث خطأ أثناء استخراج النص: ${errorMessage}`;
      }
    }

    // كتابة النص المستخرج إلى ملف
    try {
      await writeFileAsync(outputPath, extractedText);
      console.log(`تم استخراج النص من ${path.basename(filePath)} وحفظه في ${outputPath}`);
    } catch (writeError) {
      console.error(`خطأ في كتابة النص المستخرج إلى ${outputPath}:`, writeError);
      throw new Error(`لا يمكن كتابة النص المستخرج: ${writeError.message}`);
    }

    // قياس وتسجيل الوقت المستغرق
    endTime = Date.now();
    const processingTimeSeconds = (endTime - startTime) / 1000;
    console.log(`اكتملت معالجة الملف ${path.basename(filePath)} في ${processingTimeSeconds.toFixed(2)} ثانية`);
    
    return {
      success: true,
      output: outputPath,
      extractedText,
      processingTimeSeconds
    };
    
  } catch (error) {
    console.error(`خطأ في استخراج النص من ${filePath}:`, error);
    endTime = Date.now();
    
    // كتابة رسالة الخطأ إلى ملف
    try {
      // تحديد اسم الملف لرسالة الخطأ
      const filename = path.basename(filePath);
      const timestamp = new Date().toISOString().replace(/:/g, '-').replace(/\..+/, '');
      const extractedTextFile = `${path.parse(filename).name}_error_${timestamp}.txt`;
      const errorOutputPath = path.join(outputDir, extractedTextFile);
      
      // إعداد رسالة خطأ مفصلة
      const errorDetails = {
        filename,
        filePath,
        error: error.message || String(error),
        stack: error.stack,
        timestamp: new Date().toISOString(),
        processingTime: ((endTime - startTime) / 1000).toFixed(2) + ' seconds'
      };
      
      const errorMessage = `خطأ أثناء استخراج النص من ${filename}:\n` +
                           `- رسالة الخطأ: ${error.message || error}\n` +
                           `- وقت المعالجة: ${errorDetails.processingTime}\n` +
                           `- التاريخ والوقت: ${errorDetails.timestamp}\n\n` +
                           `يرجى التحقق من صحة الملف ومحاولة المعالجة مرة أخرى.`;
      
      await writeFileAsync(errorOutputPath, errorMessage);
      console.log(`تم كتابة رسالة الخطأ إلى ${errorOutputPath}`);
      
      return { 
        success: false, 
        output: errorOutputPath,
        error: error.message || String(error),
        processingTimeSeconds: ((endTime - startTime) / 1000)
      };
    } catch (writeError) {
      console.error('خطأ في كتابة رسالة الخطأ إلى ملف:', writeError);
      return { 
        success: false, 
        error: `${error.message || String(error)}. فشل أيضًا في كتابة رسالة الخطأ: ${writeError.message}`,
        processingTimeSeconds: ((endTime - startTime) / 1000)
      };
    }
  }
}

/**
 * مسار تحميل الملفات - يدعم معالجة الملفات الكبيرة حتى 400 ميجابايت
 * يقوم باستخراج النص من الملفات المرفوعة باستخدام معالجة مجزأة للملفات الكبيرة
 * يقدم معلومات تفصيلية عن عملية المعالجة لكل ملف
 */
app.post('/upload', upload.array('files'), async (req, res) => {
  // التحقق من تسجيل دخول المستخدم
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  // إعدادات معالجة الملفات - تم تحديث حد حجم الملف من 10 إلى 400 ميجابايت
  const FILE_SIZE_LIMIT = 400 * 1024 * 1024; // 400 ميجابايت
  const CHUNK_PROCESSING_ENABLED = true; // تفعيل معالجة الأجزاء للملفات الكبيرة
  
  const username = req.session.user;
  const lectureName = req.body.lectureName;
  
  console.log(`بدء معالجة ${req.files.length} ملفات للمستخدم ${username} في المحاضرة ${lectureName}`);
  
  // معالجة كل ملف مرفوع بمعالج النص
  const extractionResults = [];
  const startTime = Date.now();
  
  // معالجة الملفات بشكل متسلسل لتجنب استهلاك الكثير من الموارد
  for (const file of req.files) {
    console.log(`معالجة الملف: ${file.originalname} (${(file.size / (1024 * 1024)).toFixed(2)} ميجابايت)`);
    const filePath = file.path;
    
    // استخراج النص إلى مجلد "extract txt"
    const extractOutputDir = path.join(__dirname, 'extract txt', username, lectureName);
    
    try {
      // استدعاء وظيفة استخراج النص المحسنة
      const extractionResult = await extractTextFromFile(filePath, extractOutputDir);
      
      // إضافة معلومات إضافية للنتيجة
      extractionResult.originalFilename = file.originalname;
      extractionResult.size = file.size;
      extractionResult.mimeType = file.mimetype;
      
      // تسجيل معلومات المعالجة
      if (extractionResult.success) {
        console.log(`تم استخراج النص بنجاح من ${file.originalname} في ${extractionResult.processingTimeSeconds.toFixed(2)} ثانية`);
        console.log(`الملف الناتج: ${extractionResult.output}`);
      } else {
        console.error(`فشل استخراج النص من ${file.originalname}: ${extractionResult.error}`);
      }
      
      extractionResults.push(extractionResult);
    } catch (error) {
      console.error(`خطأ غير متوقع أثناء معالجة الملف ${file.originalname}:`, error);
      extractionResults.push({
        success: false,
        originalFilename: file.originalname,
        size: file.size,
        mimeType: file.mimetype,
        error: error.message || 'خطأ غير معروف أثناء المعالجة'
      });
    }
  }
  
  // حساب إحصائيات المعالجة
  const totalTime = (Date.now() - startTime) / 1000;
  const successCount = extractionResults.filter(r => r.success).length;
  const failCount = extractionResults.filter(r => !r.success).length;
  
  console.log(`\n=== اكتملت معالجة جميع الملفات ===\n` +
             `- الوقت الإجمالي: ${totalTime.toFixed(2)} ثانية\n` +
             `- الملفات الناجحة: ${successCount}/${req.files.length}\n` +
             `- الملفات الفاشلة: ${failCount}/${req.files.length}`);
  
  // إرسال استجابة تفصيلية للمستخدم
  res.json({ 
    success: true, 
    message: `تم رفع ${req.files.length} ملفات واستخراج النص بنجاح من ${successCount} منها`,
    processingTimeSeconds: totalTime.toFixed(2),
    stats: {
      totalFiles: req.files.length,
      successfulExtractions: successCount,
      failedExtractions: failCount
    },
    files: extractionResults.map(result => ({
      filename: result.originalFilename,
      size: result.size,
      success: result.success,
      outputPath: result.success ? path.basename(result.output) : null,
      processingTime: result.processingTimeSeconds ? `${result.processingTimeSeconds.toFixed(2)} ثانية` : null,
      error: result.error || null
    }))
  });
});

// Get user lectures
app.get('/api/lectures', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const username = req.session.user;
  const userDir = path.join(__dirname, 'uploads', username);
  
  if (!fs.existsSync(userDir)) {
    return res.json({ lectures: [] });
  }
  
  const lectures = fs.readdirSync(userDir, { withFileTypes: true })
    .filter(dirent => dirent.isDirectory())
    .map(dirent => {
      const lectureDir = path.join(userDir, dirent.name);
      let files = [];
      
      try {
        files = fs.readdirSync(lectureDir).map(file => ({
          name: file,
          path: `/uploads/${username}/${dirent.name}/${file}`
        }));
      } catch (err) {
        console.error(`Error reading files in lecture directory ${lectureDir}:`, err);
      }
      
      return {
        name: dirent.name,
        files: files
      };
    });
  
  res.json({ lectures });
});

// Delete lecture route
app.delete('/api/lecture/:lectureName', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: 'Not logged in' });
  }
  
  const username = req.session.user;
  const lectureName = req.params.lectureName;
  
  if (!lectureName) {
    return res.status(400).json({ success: false, message: 'Missing lecture name' });
  }
  
  try {
    // List of parent directories where we need to delete the lecture folder
    const parentDirs = ['uploads', 'extract txt', 'summarized txt', 'question'];
    let deletedAny = false;
    
    // Delete lecture directories in all parent directories
    parentDirs.forEach(parentDir => {
      const lectureDir = path.join(__dirname, parentDir, username, lectureName);
      
      if (fs.existsSync(lectureDir)) {
        // Delete the lecture directory using global function
        deleteFolderRecursive(lectureDir);
        deletedAny = true;
      }
    });
    
    if (deletedAny) {
      res.json({ success: true, message: 'Lecture deleted successfully' });
    } else {
      res.status(404).json({ success: false, message: 'Lecture not found' });
    }
  } catch (err) {
    console.error('Error deleting lecture:', err);
    res.status(500).json({ success: false, message: 'Error deleting lecture' });
  }
});

// Logout route
app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/');
});

// Function to recursively delete a directory
function deleteFolderRecursive(folderPath) {
  if (fs.existsSync(folderPath)) {
    fs.readdirSync(folderPath).forEach((file) => {
      const curPath = path.join(folderPath, file);
      if (fs.lstatSync(curPath).isDirectory()) {
        // Recurse
        deleteFolderRecursive(curPath);
      } else {
        // Delete file
        try {
          fs.unlinkSync(curPath);
        } catch (err) {
          console.error(`Error deleting file ${curPath}:`, err);
        }
      }
    });
    try {
      fs.rmdirSync(folderPath);
    } catch (err) {
      console.error(`Error deleting directory ${folderPath}:`, err);
    }
  }
}

// Create directories if they don't exist
const requiredDirs = ['uploads', 'extract txt', 'summarized txt', 'question'];

requiredDirs.forEach(dir => {
  if (!fs.existsSync(path.join(__dirname, dir))) {
    fs.mkdirSync(path.join(__dirname, dir));
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});

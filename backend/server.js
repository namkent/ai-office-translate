const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const dotenv = require('dotenv');
const { OpenAI } = require('openai');

// Tải các biến môi trường cấu hình từ tệp tin .env
dotenv.config();

// Tự động kiểm tra và sinh Chứng chỉ SSL nội bộ tự ký (nếu chưa tồn tại) để chạy HTTPS
const generateCerts = require('./generate-certs');
generateCerts();

// Hàm đọc và tải danh mục từ điển chuyên ngành từ tệp tin glossary.json
const getGlossary = () => {
  try {
    const glossaryPath = path.join(__dirname, 'glossary.json');
    if (fs.existsSync(glossaryPath)) {
      return JSON.parse(fs.readFileSync(glossaryPath, 'utf8'));
    }
  } catch (e) {
    console.error('[ERROR] Lỗi khi đọc file glossary.json:', e);
  }
  return null;
};

// Khởi tạo ứng dụng Express
const app = express();
const PORT = process.env.PORT || 3000;
const CLIENT_TOKEN = process.env.CLIENT_ACCESS_TOKEN || 'default-secure-token-12345';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'admin-secure-pass';

// Cấu hình CORS - Cho phép chéo nguồn để Office Add-in chạy được trên cả trình duyệt (Office Online) và ứng dụng Office Desktop
app.use(cors({
  origin: '*', 
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

// Giới hạn dung lượng yêu cầu tối đa 5MB (để đáp ứng các tài liệu Word/Excel lớn) và tự động chuyển đổi JSON body
app.use(express.json({ limit: '5mb' }));

// Cung cấp giao diện tĩnh (Single Page App) quản trị từ điển tại đường dẫn /admin
app.use('/admin', express.static(path.join(__dirname, 'admin')));

// Cấu hình Rate Limiting để ngăn chặn spam, bảo vệ API khỏi nguy cơ bị DDOS hoặc lạm dụng vượt định mức tài khoản OpenAI
const limiter = rateLimit({
  windowMs: 1 * 60 * 1000, // Chu kỳ kiểm tra 1 phút
  max: parseInt(process.env.RATE_LIMIT_MAX || '60'), // Tối đa 60 yêu cầu trên mỗi IP trong 1 phút
  message: {
    error: 'Too many requests. Please try again after a minute.',
    status: 429
  },
  standardHeaders: true, 
  legacyHeaders: false, 
});
app.use('/api/', limiter);

// Middleware xác thực bảo mật Client Token truyền lên từ COM Add-in
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1]; // Định dạng tiêu chuẩn "Bearer <token>"
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || req.ip;

  if (!token) {
    console.warn(`[SECURITY WARNING] [IP: ${ip}] Từ chối truy cập API: Thiếu Access Token.`);
    return res.status(401).json({ error: 'Truy cập bị từ chối. Thiếu Access Token.' });
  }

  if (token !== CLIENT_TOKEN) {
    console.warn(`[SECURITY WARNING] [IP: ${ip}] Từ chối truy cập API: Access Token sai.`);
    return res.status(403).json({ error: 'Access Token không hợp lệ.' });
  }

  next();
};

// Khởi tạo và cấu hình API Client kết nối tới dịch vụ OpenAI (hoặc Groq/Llama-compatible baseURL)
const openaiKey = process.env.OPENAI_KEY;
const openaiUrl = process.env.OPENAI_URL;
const isMockMode = !openaiKey || openaiKey === 'mock';

let openai = null;
if (!isMockMode) {
  const openAIConfig = {
    apiKey: openaiKey
  };
  if (openaiUrl) {
    openAIConfig.baseURL = openaiUrl;
  }
  openai = new OpenAI(openAIConfig);
}

// Middleware xác thực mật khẩu quản trị cho trang Admin
const authenticateAdmin = (req, res, next) => {
  const adminPass = req.headers['x-admin-password'];
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || req.ip;

  if (!adminPass || adminPass !== ADMIN_PASSWORD) {
    console.warn(`[SECURITY WARNING] [IP: ${ip}] Cố gắng truy cập admin thất bại.`);
    return res.status(401).json({ error: 'Truy cập bị từ chối. Mật khẩu quản trị không chính xác.' });
  }
  next();
};

// API: Lấy danh sách thuật ngữ từ điển chuyên ngành của công ty
app.get('/api/admin/glossary', authenticateAdmin, (req, res) => {
  const glossary = getGlossary() || {};
  res.json(glossary);
});

// API: Cập nhật thuật ngữ mới vào file glossary.json
app.post('/api/admin/glossary', authenticateAdmin, (req, res) => {
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || req.ip;
  try {
    const newGlossary = req.body;
    if (typeof newGlossary !== 'object') {
      return res.status(400).json({ error: 'Dữ liệu gửi lên không đúng định dạng JSON.' });
    }
    const glossaryPath = path.join(__dirname, 'glossary.json');
    fs.writeFileSync(glossaryPath, JSON.stringify(newGlossary, null, 2), 'utf8');
    res.json({ success: true });
  } catch (err) {
    console.error(`[ERROR] [IP: ${ip}] Không thể lưu glossary.json:`, err);
    res.status(500).json({ error: 'Không thể lưu từ điển chuyên ngành.' });
  }
});

// API Chính: Tiếp nhận và xử lý dịch thuật từ COM Add-in
app.post('/api/translate', authenticateToken, async (req, res) => {
  const { htmlContent, texts, targetLanguage, sourceLanguage } = req.body;
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || req.ip;

  if (!targetLanguage) {
    console.warn(`[API WARNING] [IP: ${ip}] Yêu cầu bị từ chối: Thiếu ngôn ngữ đích.`);
    return res.status(400).json({ error: 'Thiếu ngôn ngữ đích (targetLanguage).' });
  }

  // Chế độ MOCK (Chỉ dùng kiểm thử ngoại tuyến khi không cấu hình API Key để tiết kiệm chi phí)
  if (isMockMode) {
    if (htmlContent) {
      const mockHtml = htmlContent.replace(/>([^<]+)</g, (match, text) => {
        if (text.trim()) {
          return `>[${targetLanguage}] ${text}<`;
        }
        return match;
      });
      return res.json({ translatedHtml: mockHtml });
    } else if (texts && Array.isArray(texts)) {
      const mockTexts = texts.map(t => t.trim() ? `[${targetLanguage}] ${t}` : t);
      return res.json({ translatedTexts: mockTexts });
    }
    return res.status(400).json({ error: 'Cần gửi htmlContent hoặc mảng texts để dịch.' });
  }

  try {
    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';

    // TRƯỜNG HỢP 1: Dịch tài liệu có định dạng HTML (Dành cho Microsoft Word)
    if (htmlContent !== undefined) {
      if (typeof htmlContent !== 'string') {
        return res.status(400).json({ error: 'htmlContent phải là dạng string.' });
      }

      // MÃ HÓA BẢO VỆ BULLET: Tìm các thẻ span sử dụng font Symbol/Wingdings để hiển thị bullet.
      // Di chuyển ký tự bullet (ví dụ: chữ 'l') vào thuộc tính tạm 'data-bullet-temp' để tránh AI dịch hoặc xóa mất.
      const bulletRegex = /<span style=(['"])([^'"]*font-family:(Symbol|Wingdings|Webdings)[^'"]*)\1>([^<]+)/gi;
      let maskedHtml = htmlContent.replace(bulletRegex, '<span style=$1$2$1 data-bullet-temp=$1$4$1>');

      // Đọc glossary chuyên ngành để áp đặt quy tắc dịch thuật vào Prompt hệ thống
      const glossary = getGlossary();
      let glossaryRule = "";
      if (glossary && Object.keys(glossary).length > 0) {
        glossaryRule = `\nCRITICAL TERMINOLOGY GLOSSARY (You MUST translate these exact terms accordingly if present in the text):\n` +
          Object.entries(glossary).map(([src, tgt]) => `- "${src}" -> "${tgt}"`).join('\n') + "\n";
      }

      const srcLang = sourceLanguage && sourceLanguage !== 'Auto Detect' ? sourceLanguage : 'its current language';
      
      // Xây dựng Prompt dịch HTML nghiêm ngặt để bảo toàn cấu trúc thẻ, chỉ dịch text hiển thị
      const systemPrompt = `You are a professional translator. Translate the given HTML content from ${srcLang} into the target language: ${targetLanguage}.
${glossaryRule}
CRITICAL RULES:
1. Translate ONLY the visible text contents. Do NOT translate or modify any HTML tags, tag names, attributes (e.g. style, class, id, href, val, lang, face), or HTML entities.
2. Maintain the exact tag structure and nesting. Do not merge, split, or omit any tags.
3. Keep all punctuation and spacing surrounding the text within the tags as natural as possible in the translated version.
4. Do NOT enclose your output in markdown code blocks (like \`\`\`html ... \`\`\`). Output only the raw translated HTML.
5. If the input is empty or contains no translatable text, return the input exactly as is.
6. Do NOT modify, translate, or replace bullet characters (for example, characters inside tags with style 'font-family:Symbol', commonly the letter 'l' or symbols like '&middot;'). Keep them exactly as they are in the original HTML.
7. MANDATORY TARGET LANGUAGE: You MUST translate the text into the target language: ${targetLanguage}. Under no circumstances should you translate it into any other language (such as English). Even if the text is surrounded by HTML tags with language attributes like 'lang="en-US"' or 'lang="en"', ignore those attributes for the translation direction and translate the visible text to ${targetLanguage}.`;

      const response = await openai.chat.completions.create({
        model: model,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: maskedHtml } 
        ],
        temperature: 0.3,
      });

      let translatedHtml = response.choices[0].message.content.trim();
      
      // Dọn dẹp khối mã markdown (```html) nếu mô hình LLM vô tình trả về
      if (translatedHtml.startsWith('```html')) {
        translatedHtml = translatedHtml.replace(/^```html\s*/i, '').replace(/\s*```$/i, '');
      } else if (translatedHtml.startsWith('```')) {
        translatedHtml = translatedHtml.replace(/^```\s*/i, '').replace(/\s*```$/i, '');
      }

      // GIẢI MÃ KHÔI PHỤC BULLET: Đưa các ký tự bullet nguyên bản từ thuộc tính tạm trở lại vùng hiển thị của Word
      const restoreRegex = /<span style=(['"])([^'"]*font-family:(Symbol|Wingdings|Webdings)[^'"]*)\1 data-bullet-temp=(['"])([^'"]*)\4>/gi;
      let restoredHtml = translatedHtml.replace(restoreRegex, '<span style=$1$2$1>$5');

      // CHUẨN HÓA KÝ TỰ ĐẶC BIỆT: Thay các khoảng trắng bất thường do AI tự ý sinh ra về thực thể &nbsp; tiêu chuẩn
      // Điều này ngăn ngừa Word bị lỗi vẽ thành các ô vuông lỗi font trên màn hình
      let sanitizedHtml = restoredHtml
        .replace(/\u00A0/g, '&nbsp;')          
        .replace(/[\u200B-\u200D\uFEFF]/g, '')   
        .replace(/\t/g, '    ');               

      return res.json({ translatedHtml: sanitizedHtml });
    }

    // TRƯỜNG HỢP 2: Dịch mảng chuỗi văn bản (Dành cho Excel Cells và PowerPoint TextBoxes để tối ưu hóa token)
    if (texts !== undefined) {
      if (!Array.isArray(texts)) {
        return res.status(400).json({ error: 'texts phải là một mảng strings.' });
      }

      const glossary = getGlossary();
      let glossaryRule = "";
      if (glossary && Object.keys(glossary).length > 0) {
        glossaryRule = `\nCRITICAL TERMINOLOGY GLOSSARY (You MUST translate these exact terms accordingly if present in the text):\n` +
          Object.entries(glossary).map(([src, tgt]) => `- "${src}" -> "${tgt}"`).join('\n') + "\n";
      }

      const srcLang = sourceLanguage && sourceLanguage !== 'Auto Detect' ? sourceLanguage : 'their current language';
      
      // Xây dựng Prompt ép cấu trúc định dạng JSON kết quả dịch
      const systemPrompt = `You are an expert translator. Translate the array of text strings from ${srcLang} into the target language: ${targetLanguage}.
${glossaryRule}
CRITICAL RULES:
1. The output MUST be a valid JSON object containing a "translations" key which holds the array of translated strings, matching the input array length exactly.
   Example Input:
   {
     "texts": ["Hello", "World"]
   }
   Example Output:
   {
     "translations": ["Xin chào", "Thế giới"]
   }
2. Do NOT change the order or count of items. Keep empty strings or formatting exactly as is.
3. Translate each string accurately. If a string is empty, contains only whitespace, or contains only numbers/formulas, keep it exactly as is in the output array.
4. Return ONLY the raw JSON object. Do NOT wrap it in markdown code blocks or write explanations.
5. MANDATORY TARGET LANGUAGE: You MUST translate every text string into the target language: ${targetLanguage}. Do not translate into English or any other language unless the target language is English.`;

      const response = await openai.chat.completions.create({
        model: model,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: JSON.stringify({ texts: texts }) }
        ],
        response_format: { type: "json_object" }, // Kích hoạt chế độ xác thực Schema JSON của OpenAI/Groq
        temperature: 0.3,
      });

      const responseContent = response.choices[0].message.content.trim();
      let translatedTexts;
      
      try {
        const parsed = JSON.parse(responseContent);
        if (Array.isArray(parsed)) {
          translatedTexts = parsed;
        } else if (parsed.translations && Array.isArray(parsed.translations)) {
          translatedTexts = parsed.translations;
        } else {
          const firstKey = Object.keys(parsed)[0];
          if (Array.isArray(parsed[firstKey])) {
            translatedTexts = parsed[firstKey];
          } else {
            throw new Error('Không tìm thấy mảng dịch thuật.');
          }
        }
      } catch (err) {
        console.error(`[ERROR] [IP: ${ip}] Lỗi phân giải JSON từ OpenAI:`, responseContent, err);
        return res.status(500).json({ error: 'Lỗi cấu trúc dữ liệu JSON từ OpenAI.', rawResponse: responseContent });
      }

      if (translatedTexts.length !== texts.length) {
        console.warn(`[WARNING] [IP: ${ip}] Mảng dịch bị lệch kích thước. Client: ${texts.length}, AI: ${translatedTexts.length}`);
      }

      return res.json({ translatedTexts });
    }

    return res.status(400).json({ error: 'Request không hợp lệ. Cần truyền htmlContent hoặc mảng texts.' });

  } catch (error) {
    console.error(`[ERROR] [IP: ${ip}] Lỗi xử lý dịch thuật API:`, error);
    return res.status(500).json({ error: 'Lỗi server khi gọi API dịch thuật: ' + error.message });
  }
});

// Khởi động HTTPS Server với chứng chỉ bảo mật tự ký nội bộ
try {
  const key = fs.readFileSync(path.join(__dirname, 'key.pem'), 'utf8');
  const cert = fs.readFileSync(path.join(__dirname, 'cert.pem'), 'utf8');

  https.createServer({ key, cert }, app).listen(PORT, () => {
    console.log(`=================================================`);
    console.log(`AI Translate Backend Server is running at:`);
    console.log(`HTTPS Server: https://localhost:${PORT}`);
    console.log(`Add-in Path:  https://localhost:${PORT}/addon/taskpane.html`);
    console.log(`API Translate: https://localhost:${PORT}/api/translate`);
    console.log(`Mode:         ${isMockMode ? 'MOCK MODE (No API Key)' : 'LIVE MODE (OpenAI)'}`);
    console.log(`=================================================`);
  });
} catch (err) {
  console.error('[ERROR] Không thể khởi động HTTPS server do thiếu cert.pem/key.pem.', err);
  process.exit(1);
}

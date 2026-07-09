const express = require('express');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// Allow requests from anywhere
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Cache-Control, Pragma');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Setup memory storage for incoming file processing pipeline
const upload = multer({ storage: multer.memoryStorage() });

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'Email B backend is running!' });
});

/* ══ NEW: MASS DATA CONFIGURATION IMPORT PIPELINE (.TXT, .CSV, .XLSX) ══ */
app.post('/upload-pipeline', upload.single('accountsFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded' });
    }

    const filename = req.file.originalname.toLowerCase();
    const accounts = [];

    // Process Spreadsheets (both Excel .xlsx and comma/tab-separated .csv)
    if (filename.endsWith('.xlsx') || filename.endsWith('.csv')) {
      const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

      for (const row of jsonRows) {
        if (!row || row.length === 0) continue;

        // Condition A: Data is all crammed into a single cell using pipe characters
        if (row.length === 1 && typeof row[0] === 'string' && row[0].includes('|')) {
          const line = row[0].trim();
          const parts = line.split('|');
          if (parts.length >= 4) {
            accounts.push({ email: parts[0], password: parts[1], refresh_token: parts[2], client_id: parts[3], raw: line });
          }
        } else if (row.length >= 4) {
          // Condition B: Data is beautifully separated across columns (Col A: Email, Col B: Pass, etc.)
          const email = String(row[0]).trim();
          const password = String(row[1]).trim();
          const refresh_token = String(row[2]).trim();
          const client_id = String(row[3]).trim();
          if (email && password && refresh_token && client_id) {
            accounts.push({
              email,
              password,
              refresh_token,
              client_id,
              raw: `${email}|${password}|${refresh_token}|${client_id}`
            });
          }
        }
      }
      return res.json({ success: true, accounts });
    } else {
      // Process standard line-break text files (.txt)
      const textContent = req.file.buffer.toString('utf-8');
      const lines = textContent.split(/\r?\n/);
      
      lines.forEach(line => {
        const cleaned = line.trim();
        if (!cleaned) return;
        const parts = cleaned.split('|');
        if (parts.length >= 4) {
          accounts.push({ email: parts[0], password: parts[1], refresh_token: parts[2], client_id: parts[3], raw: cleaned });
        }
      });
      return res.json({ success: true, accounts });
    }
  } catch (error) {
    return res.status(500).json({ success: false, message: error.message });
  }
});

/* ══ NEW: SECURE MICROSOFT REST API OAUTH & INBOX OTP READER ══ */
app.post('/read-hotmail-code', async (req, res) => {
  const { refresh_token, client_id } = req.body;
  if (!refresh_token || !client_id) {
    return res.status(400).json({ success: false, message: 'Missing credential keys' });
  }

  try {
    // 1. Swap the refresh token for a live secure temporary session Access Token
    const tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    const params = new URLSearchParams();
    params.append('client_id', client_id);
    params.append('scope', 'https://graph.microsoft.com/Mail.Read');
    params.append('refresh_token', refresh_token);
    params.append('grant_type', 'refresh_token');

    const tokenRes = await axios.post(tokenUrl, params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const accessToken = tokenRes.data.access_token;

    // 2. Request last 5 items out of user inbox container
    const mailUrl = 'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,body';
    const mailRes = await axios.get(mailUrl, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    const messages = mailRes.data.value || [];
    let extractedCode = null;

    // 3. Isolated internal verification parser logic mapping string formats
    function findCode(text) {
      if (!text) return null;
      const subjectStartMatch = text.match(/^(\d{4,8})\s+(?:is your|ist dein|is the|es tu)/im);
      if (subjectStartMatch) return subjectStartMatch[1];

      const inlineMatch = text.match(/(?:code|pin|otp)[^\d]{0,20}(\d{4,8})(?:\s|$)/i);
      if (inlineMatch) return inlineMatch[1];

      const genericMatch = text.match(/\b\d{4,8}\b/);
      if (genericMatch) return genericMatch[0];
      return null;
    }

    for (const msg of messages) {
      const subject = msg.subject || '';
      const bodyText = msg.body?.content || '';
      const cleanBody = bodyText.replace(/<[^>]*>/g, ' '); // Strip HTML rendering tags
      
      const parsedOutput = findCode(subject + '\n' + cleanBody);
      if (parsedOutput) {
        extractedCode = parsedOutput;
        break;
      }
    }

    if (extractedCode) {
      return res.json({ success: true, code: extractedCode });
    } else {
      return res.json({ success: false, message: 'No dynamic code matching triggers located inside messages.' });
    }
  } catch (error) {
    const errMsg = error.response?.data?.error_description || error.message;
    return res.status(500).json({ success: false, message: 'Authentication routine error: ' + errMsg });
  }
});

/* ══ UNTOUCHED LEGACY INBOX IMAP COEXISTENCE LAYER ════════════ */
app.post('/emails', (req, res) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.set('Pragma', 'no-cache');

  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required' });
  }

  let imapConfig = {
    user: email,
    password: password,
    tls: true,
    tlsOptions: { rejectUnauthorized: false },
    connTimeout: 15000,
    authTimeout: 15000,
  };

  if (email.includes('@outlook') || email.includes('@hotmail') || email.includes('@live')) {
    imapConfig.host = 'outlook.office365.com';
    imapConfig.port = 993;
  } else if (email.includes('@gmail')) {
    imapConfig.host = 'imap.gmail.com';
    imapConfig.port = 993;
  } else if (email.includes('@yahoo')) {
    imapConfig.host = 'imap.mail.yahoo.com';
    imapConfig.port = 993;
  } else if (email.includes('@gmx')) {
    imapConfig.host = 'imap.gmx.com';
    imapConfig.port = 993;
  } else if (email.includes('@zoho')) {
    imapConfig.host = email.includes('.eu') ? 'imap.zoho.eu' : 'imap.zoho.com';
    imapConfig.port = 993;
  } else {
    const domain = email.split('@')[1];
    imapConfig.host = 'imap.' + domain;
    imapConfig.port = 993;
  }

  const imap = new Imap(imapConfig);
  let responded = false;

  function safeRespond(data) {
    if (!responded) {
      responded = true;
      res.json(data);
    }
  }

  imap.once('ready', () => {
    imap.getBoxes((err, boxes) => {
      if (err) {
        imap.end();
        return safeRespond({ error: 'Could not list folders: ' + err.message });
      }

      const allFolders = [];
      function flatten(obj, prefix) {
        for (const name in obj) {
          const delim = obj[name].delimiter || '/';
          const full = prefix ? prefix + delim + name : name;
          allFolders.push(full);
          if (obj[name].children) flatten(obj[name].children, full);
        }
      }
      flatten(boxes, '');

      let targetFolder = 'INBOX';
      const notifMatch = allFolders.find(f =>
        f.toLowerCase().includes('notification') ||
        f.toLowerCase().includes('powiadomien')
      );
      if (notifMatch) targetFolder = notifMatch;

      imap.openBox(targetFolder, false, (err, box) => {
        if (err) {
          imap.openBox('INBOX', false, (err2, box2) => {
            if (err2 || !box2) {
              imap.end();
              return safeRespond({ error: 'Could not open inbox: ' + (err2 ? err2.message : 'unknown') });
            }
            fetchFromBox(box2);
          });
          return;
        }
        fetchFromBox(box);
      });

      function fetchFromBox(box) {
        if (!box || box.messages.total === 0) {
          imap.end();
          return safeRespond({ emails: [] });
        }

        const total = box.messages.total;
        const start = Math.max(1, total - 1);
        const fetcher = imap.seq.fetch(`${start}:${total}`, {
          bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)', 'TEXT'],
          struct: true,
          markSeen: false,
        });

        const collected = [];
        const pending = [];

        fetcher.on('message', (msg, seqno) => {
          const emailData = { id: seqno, from: '', subject: '', date: '', body: '', isRead: false };

          msg.on('body', (stream, info) => {
            const p = new Promise((resolve) => {
              simpleParser(stream, (err, parsed) => {
                if (!err) {
                  if (info.which.includes('HEADER')) {
                    emailData.from = parsed.from?.text || '';
                    emailData.subject = parsed.subject || '';
                    emailData.date = parsed.date?.toISOString() || new Date().toISOString();
                  } else {
                    emailData.body = parsed.text || '';
                  }
                }
                resolve();
              });
            });
            pending.push(p);
          });

          msg.once('attributes', (attrs) => {
            emailData.isRead = attrs.flags && attrs.flags.includes('\\Seen');
          });

          msg.once('end', () => collected.push(emailData));
        });

        fetcher.once('error', (err) => {
          imap.end();
          safeRespond({ error: 'Fetch error: ' + err.message });
        });

        fetcher.once('end', () => {
          Promise.all(pending).then(() => {
            imap.end();
            collected.sort((a, b) => new Date(b.date) - new Date(a.date));
            safeRespond({ emails: collected });
          });
        });
      }
    });
  });

  imap.once('error', (err) => {
    safeRespond({ error: 'Connection failed: ' + err.message });
  });

  imap.connect();
});

app.listen(PORT, () => {
  console.log(`Email B backend running on port ${PORT}`);

  // Self-ping every 10 minutes to prevent sleeping
  setInterval(() => {
    const https = require('https');
    https.get('https://emailb-backend.onrender.com/', (res) => {
      console.log(`Self-ping: ${res.statusCode}`);
    }).on('error', (e) => {
      console.log(`Self-ping failed: ${e.message}`);
    });
  }, 10 * 60 * 1000);
});

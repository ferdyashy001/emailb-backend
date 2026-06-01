const express = require('express');
const Imap = require('imap');
const { simpleParser } = require('mailparser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// Allow requests from anywhere
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'Email B backend is running!' });
});

// Fetch emails via IMAP
app.post('/emails', (req, res) => {
  const { email, password } = req.body;

  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required' });
  }

  // Detect provider from email address
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
    // Generic fallback — tries imap.yourdomain.com
    const domain = email.split('@')[1];
    imapConfig.host = 'imap.' + domain;
    imapConfig.port = 993;
  }

  const imap = new Imap(imapConfig);
  const emails = [];

  imap.once('ready', () => {
    // Search all folders including Notification, Spam, etc.
    imap.getBoxes((err, boxes) => {
      if (err) {
        imap.end();
        return res.status(500).json({ error: 'Could not get folders: ' + err.message });
      }

      // Flatten all folder names
      const folderNames = [];
      function getFolders(obj, prefix) {
        for (const name in obj) {
          const fullName = prefix ? prefix + obj[name].delimiter + name : name;
          folderNames.push(fullName);
          if (obj[name].children) getFolders(obj[name].children, fullName);
        }
      }
      getFolders(boxes, '');

      // Always include INBOX plus common folders
      const foldersToSearch = ['INBOX'];
      const extras = ['Notification', 'Spam', 'Junk', 'Bulk Mail', 'Bulk', 'Social', 'Promotions', 'Updates', 'Forums'];
      extras.forEach(f => {
        const match = folderNames.find(n => n.toLowerCase().includes(f.toLowerCase()));
        if (match && !foldersToSearch.includes(match)) foldersToSearch.push(match);
      });

      const allEmails = [];
      let folderIndex = 0;

      function searchNextFolder() {
        if (folderIndex >= foldersToSearch.length) {
          imap.end();
          allEmails.sort((a, b) => new Date(b.date) - new Date(a.date));
          return res.json({ emails: allEmails.slice(0, 50) });
        }

        const folder = foldersToSearch[folderIndex++];
        imap.openBox(folder, false, (err, box) => {
          if (err || !box || box.messages.total === 0) {
            return searchNextFolder();
          }

          const total = box.messages.total;
          const start = Math.max(1, total - 29);
          const fetch = imap.seq.fetch(`${start}:${total}`, {
            bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)', 'TEXT'],
            struct: true,
            markSeen: false,
          });

          const pending = [];

          fetch.on('message', (msg, seqno) => {
            const emailData = { id: `${folder}_${seqno}`, from: '', subject: '', date: '', body: '', isRead: false };

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

            msg.once('end', () => {
              allEmails.push(emailData);
            });
          });

          fetch.once('error', () => searchNextFolder());

          fetch.once('end', () => {
            Promise.all(pending).then(() => searchNextFolder());
          });
        });
      }

      searchNextFolder();
    });
  });

  imap.once('error', (err) => {
    res.status(500).json({ error: 'Connection failed: ' + err.message });
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

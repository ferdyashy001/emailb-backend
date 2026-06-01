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
    imap.openBox('INBOX', false, (err, box) => {
      if (err) {
        imap.end();
        return res.status(500).json({ error: 'Could not open inbox: ' + err.message });
      }

      const total = box.messages.total;
      if (total === 0) {
        imap.end();
        return res.json({ emails: [] });
      }

      const start = Math.max(1, total - 49);
      const fetch = imap.seq.fetch(`${start}:${total}`, {
        bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)', 'TEXT'],
        struct: true,
        markSeen: false,
      });

      const pending = [];

      fetch.on('message', (msg, seqno) => {
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

        msg.once('end', () => {
          emails.push(emailData);
        });
      });

      fetch.once('error', (err) => {
        imap.end();
        res.status(500).json({ error: 'Fetch error: ' + err.message });
      });

      fetch.once('end', () => {
        Promise.all(pending).then(() => {
          imap.end();
          emails.sort((a, b) => new Date(b.date) - new Date(a.date));
          res.json({ emails: emails.slice(0, 50) });
        });
      });
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

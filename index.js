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
    // First get all folder names so we can find the right one
    imap.getBoxes((err, boxes) => {
      if (err) {
        imap.end();
        return safeRespond({ error: 'Could not list folders: ' + err.message });
      }

      // Flatten all folders
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

      // Find Notification folder (works in any language)
      // Priority: Notification > INBOX
      let targetFolder = 'INBOX';
      const notifMatch = allFolders.find(f =>
        f.toLowerCase().includes('notification') ||
        f.toLowerCase().includes('powiadomien') // Polish
      );
      if (notifMatch) targetFolder = notifMatch;

      imap.openBox(targetFolder, false, (err, box) => {
        if (err) {
          // Try INBOX as fallback
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
        // Fetch last 2 emails only
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

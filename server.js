import express from 'express';
import path from 'path';
import fs from 'fs';
import multer from 'multer';
import libreoffice from 'libreoffice-convert';
import { execFile } from 'child_process';
import os from 'os';

// Helper wrapper: libreoffice.convert expects a callback as the 4th arg
// (convert(document, format, filter, callback)). If we want a Promise API
// and no filter, call with `null` for filter and provide a callback.
const convertAsync = (document, format, fileName) => {
  return new Promise((resolve, reject) => {
    try {
      // Prefer convertWithOptions so we can pass a filename (helps soffice detect format)
      if (typeof libreoffice.convertWithOptions === 'function') {
        const options = { fileName: fileName || 'source' };
        // If an explicit soffice exe path is provided, set the exec cwd to its directory
        // to help LibreOffice find its platform independent libraries on Windows.
        try {
          const sofficePath = process.env.LIBRE_OFFICE_EXE || '';
          if (sofficePath) {
            const sofficeDir = path.dirname(sofficePath);
            options.execOptions = { cwd: sofficeDir };
          }
        } catch (e) {
          // ignore errors resolving path
        }

        return libreoffice.convertWithOptions(document, format, null, options, (err, done) => {
          if (err) return reject(err);
          resolve(done);
        });
      }

      // Fallback to convert (which may not accept options)
      libreoffice.convert(document, format, null, (err, done) => {
        if (err) return reject(err);
        resolve(done);
      });
    } catch (err) {
      reject(err);
    }
  });
};


const app = express();
const PORT = 3000;
const upload = multer({ dest: 'uploads/' });

app.set('view engine', 'ejs');
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.get('/', (req, res) => {
    res.render('index');
});



// Conversion endpoint: accepts a single file field named 'file' and returns a PDF
app.post('/convert', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded (field name should be `file`)'});
  }

  const tempPath = req.file.path;
  try {
  const input = fs.readFileSync(tempPath);
  // Use format 'pdf' (no leading dot) expected by libreoffice-convert
  const outputFormat = 'pdf';
  // Use the original filename (with extension) to help LibreOffice detect the input format
  const original = req.file.originalname || `source${path.extname(req.file.originalname || '')}`;
    let pdfBuf;

    // On Windows, calling the soffice executable directly is often more reliable than
    // going through the wrapper which uses temp install dirs. Use direct execFile here.
    if (process.platform === 'win32') {
      const sofficeCmd = process.env.LIBRE_OFFICE_EXE || 'soffice';
      const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lo-out-'));
  // Copy uploaded file into the outDir with the original filename (helps soffice detect type)
  const inPath = path.join(outDir, original);
  fs.copyFileSync(tempPath, inPath);
  const args = ['--headless', '--convert-to', 'pdf', '--outdir', outDir, inPath];
      const execOpts = {};
      try {
        if (process.env.LIBRE_OFFICE_EXE) execOpts.cwd = path.dirname(process.env.LIBRE_OFFICE_EXE);
      } catch (e) {}

      await new Promise((resolve, reject) => {
        execFile(sofficeCmd, args, execOpts, (err, stdout, stderr) => {
          if (err) return reject(new Error(`${err.message}\n${stderr || ''}\n${stdout || ''}`));
          // log stdout/stderr for debugging
          if (stdout) console.log('soffice stdout:', stdout);
          if (stderr) console.log('soffice stderr:', stderr);
          resolve();
        });
      });

      const files = fs.readdirSync(outDir);
      console.log('Outdir files:', files);
      const baseName = path.basename(original, path.extname(original) || '');
      const outPath = path.join(outDir, `${baseName}.pdf`);
      if (!fs.existsSync(outPath)) throw new Error(`Converted file not found; outdir files: ${files.join(', ')}`);
      pdfBuf = fs.readFileSync(outPath);
      // cleanup output file
      try { fs.unlinkSync(outPath); } catch (e) {}
      try { fs.rmdirSync(outDir); } catch (e) {}
    } else {
      pdfBuf = await convertAsync(input, outputFormat, original);
    }

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': 'attachment; filename="output.pdf"'
    });
    return res.send(pdfBuf);
  } catch (err) {
    console.error('Conversion error:', err && err.message ? err.message : err);
    return res.status(500).json({ error: (err && err.message) || String(err) });
  } finally {
    // cleanup uploaded file
    fs.unlink(tempPath, () => {});
  }
});

// Export app for testing. Only start the listener when not running tests.
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
  });
}

export default app;
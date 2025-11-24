import request from 'supertest';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import app from '../server.js';

(async () => {
    try {
    // Quick check: ensure soffice (LibreOffice) is available. Try several options and
    // set process.env.LIBRE_OFFICE_EXE to a discovered exe path so the server picks it up.
    const candidates = [];
    if (process.env.LIBRE_OFFICE_EXE) candidates.push(process.env.LIBRE_OFFICE_EXE);
    candidates.push('soffice');
    candidates.push('C:\\Program Files\\LibreOffice\\program\\soffice.exe');
    candidates.push('C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe');

    let found = false;
    for (const candidate of candidates) {
      // eslint-disable-next-line no-await-in-loop
      const ok = await new Promise((resolve) => {
        // If candidate contains spaces (Windows path), quote it in the command
        const cmd = candidate === 'soffice' ? 'soffice --version' : `"${candidate}" --version`;
        exec(cmd, (err, stdout, stderr) => resolve(!err));
      });
      if (ok) {
        // If we discovered an explicit exe path, set LIBRE_OFFICE_EXE so the library uses it
        if (candidate !== 'soffice') process.env.LIBRE_OFFICE_EXE = candidate;
        found = true;
        break;
      }
    }

    if (!found) {
      console.log('Skipping conversion test: LibreOffice `soffice` not found or not usable. Install LibreOffice or set LIBRE_OFFICE_EXE.');
      process.exit(0);
    }
    const fixturePath = path.join(process.cwd(), 'tests', 'fixtures', 'sample.txt');
    if (!fs.existsSync(fixturePath)) {
      console.error('Fixture missing:', fixturePath);
      process.exit(2);
    }

    console.log('Uploading fixture to /convert...');
    const res = await request(app)
      .post('/convert')
      .attach('file', fixturePath);

    if (res.status !== 200) {
      console.error('Conversion failed. status:', res.status);
      console.error('body:', res.body || res.text || 'no body');
      // If LibreOffice is not installed, the library typically errors here.
      process.exit(3);
    }

    const ct = res.headers['content-type'] || '';
    if (!ct.includes('application/pdf')) {
      console.error('Unexpected Content-Type:', ct);
      process.exit(4);
    }

    // Save output for manual inspection
    const outPath = path.join(process.cwd(), 'tests', 'fixtures', 'out.pdf');
    fs.writeFileSync(outPath, res.body);
    console.log('Test passed — PDF written to', outPath);
    process.exit(0);
  } catch (err) {
    console.error('Test script error:', err && err.message ? err.message : err);
    process.exit(1);
  }
})();

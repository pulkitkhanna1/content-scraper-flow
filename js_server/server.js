/**
 * server.js — Novel Scraper (Node.js/Express)
 *
 * Run locally:   node server.js
 * Deploy:        works on Railway, Render, Fly.io, any Node host
 *
 * Requires Node 18+ (uses built-in fetch).
 */

import express        from 'express';
import { EventEmitter } from 'events';
import { randomBytes }  from 'crypto';
import { URL }          from 'url';
import fs               from 'fs';
import path             from 'path';
import { fileURLToPath } from 'url';
import { exec }          from 'child_process';
import * as cheerio      from 'cheerio';
import { Document, Packer, Paragraph, HeadingLevel, TextRun } from 'docx';
// googleapis no longer needed — uploads go via Apps Script

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Config ─────────────────────────────────────────────────────────────────────

const PORT          = parseInt(process.env.PORT || '7799');
const IS_SERVER     = !!(process.env.RAILWAY_ENVIRONMENT || process.env.RENDER || process.env.SERVER_MODE);
const HOME          = process.env.HOME || process.env.USERPROFILE || '/tmp';
const DEFAULT_OUT   = process.env.OUT_DIR || (IS_SERVER ? '/tmp/novels' : path.join(HOME, 'Downloads', 'Novels'));
const GDRIVE_FOLDER    = '1UyCUOcPTQLGSkII4DoEPd-_gKGZwLO9E';
const APPS_SCRIPT_URL  = process.env.APPS_SCRIPT_URL || 'https://script.google.com/macros/s/AKfycbxVW8CYPPopNWZvM3IKEuwjYqrWSzs_6tznwMsQj6gjT7h1antrYwxIzSNQhcfYHU0u/exec';
const CHAPTERS_PER_FILE = 100;

const app = express();
app.use(express.json());

// ── HTTP helpers ───────────────────────────────────────────────────────────────

const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Accept-Language': 'en-US,en;q=0.9',
  'Accept': 'text/html,application/xhtml+xml,*/*;q=0.8',
  'Referer': 'https://www.google.com/',
};

async function fetchHtml(url, timeoutMs = 15000) {
  const ctrl  = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, { headers: HEADERS, signal: ctrl.signal });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return cheerio.load(await res.text());
  } finally {
    clearTimeout(timer);
  }
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Platform registry ──────────────────────────────────────────────────────────

const PLATFORMS = {
  'webnovel.com':     { platform: 'webnovel',     lang: 'en', needs_translation: false, scraper: 'webnovel_content_uc.py',      needs_browser: true,  needs_login: true  },
  '69shuba.com':      { platform: '69shuba',      lang: 'zh', needs_translation: true,  scraper: '69shuba_new.py',              needs_browser: true,  needs_login: false },
  '69shu.com':        { platform: '69shuba',      lang: 'zh', needs_translation: true,  scraper: '69shuba_new.py',              needs_browser: true,  needs_login: false },
  'novelbin.com':     { platform: 'novelbin',     lang: 'en', needs_translation: false, scraper: 'novelbin.py',                needs_browser: false, needs_login: false },
  'freewebnovel.com': { platform: 'freewebnovel',  lang: 'en', needs_translation: false, scraper: 'freewebnovel/script_new.py', needs_browser: true,  needs_login: false },
  'babelnovel.com':   { platform: 'babelnovel',   lang: 'en', needs_translation: false, scraper: 'babelnovel_content.py',      needs_browser: true,  needs_login: true  },
  'royalroad.com':    { platform: 'royalroad',    lang: 'en', needs_translation: false, scraper: 'royalroad_content.py',       needs_browser: false, needs_login: false },
  'tapas.io':         { platform: 'tapas',        lang: 'en', needs_translation: false, scraper: 'tapas_content_fixed.py',     needs_browser: true,  needs_login: true  },
  'wuxiaworld.com':   { platform: 'wuxiaworld',   lang: 'en', needs_translation: false, scraper: 'wuxiaworld_next.py',         needs_browser: true,  needs_login: true  },
  'wattpad.com':      { platform: 'wattpad',      lang: 'en', needs_translation: false, scraper: 'wattpad_content.py',         needs_browser: true,  needs_login: false },
  'kakao.com':        { platform: 'kakao',        lang: 'ko', needs_translation: true,  scraper: 'kakao_content.py',           needs_browser: true,  needs_login: true  },
  'kakaopage.com':    { platform: 'kakao',        lang: 'ko', needs_translation: true,  scraper: 'kakao_content.py',           needs_browser: true,  needs_login: true  },
  'qidian.com':       { platform: 'qidian',       lang: 'zh', needs_translation: true,  scraper: 'qdmm_content_new.py',        needs_browser: true,  needs_login: false },
  'qdmm.com':         { platform: 'qdmm',         lang: 'zh', needs_translation: true,  scraper: 'qdmm_content_new.py',        needs_browser: true,  needs_login: false },
  'hengyan.com':      { platform: 'hengyan',      lang: 'zh', needs_translation: true,  scraper: 'hengyan/s.py',               needs_browser: true,  needs_login: false },
  'ihengyan.com':     { platform: 'hengyan',      lang: 'zh', needs_translation: true,  scraper: 'hengyan/s.py',               needs_browser: true,  needs_login: false },
  'novelupdates.com': { platform: 'novelupdates', lang: 'en', needs_translation: false, scraper: null,                         needs_browser: false, needs_login: false },
};

function detectPlatform(url) {
  try {
    let host = new URL(url).hostname.toLowerCase().replace(/^www\./, '');
    for (const [domain, cfg] of Object.entries(PLATFORMS)) {
      if (host === domain || host.endsWith('.' + domain))
        return { ...cfg, canonical_url: url };
    }
  } catch {}
  return null;
}

// ── Metadata extractors ────────────────────────────────────────────────────────

async function extractRoyalroad(url) {
  const m    = url.match(/\/fiction\/(\d+)/);
  const base = m ? `https://www.royalroad.com/fiction/${m[1]}` : url;
  const $    = await fetchHtml(base);

  const title  = $('.fic-header h1, h1.font-white, h1').first().text().trim();
  const author = $('[property="author"], .fic-header a.a-white').first().text().trim();
  const desc   = $('.description p, .description .hidden-content').first().text().trim();

  const rows = $('table#chapters tbody tr');
  const chapter_count = rows.length;

  let firstUrl = '';
  const btn = $('a.btn-read-now').first();
  if (btn.length) {
    const href = btn.attr('href') || '';
    firstUrl = href.startsWith('http') ? href : `https://www.royalroad.com${href}`;
  }
  if (!firstUrl && rows.length) {
    const href = rows.first().find('td a[href*="/chapter/"]').attr('href') || '';
    firstUrl = href.startsWith('http') ? href : `https://www.royalroad.com${href}`;
  }
  if (!firstUrl) {
    const any = $('a[href*="/chapter/"]').first().attr('href') || '';
    firstUrl = any.startsWith('http') ? any : `https://www.royalroad.com${any}`;
  }

  return { title, author, chapter_count, desc, scraper_args: { url: firstUrl || url } };
}

async function extractNovelbin(url) {
  const $ = await fetchHtml(url);
  const title  = $('h3.title, h1.title, .book-info h3').first().text().trim();
  const author = $('.author a, .info-item a').first().text().trim();
  const desc   = $('#tab-description p, .desc-text p').first().text().trim();

  let chapter_count = 0;
  const m = ($('.chapter-count, .l-chapter a').first().text() || '').match(/(\d+)/);
  if (m) chapter_count = parseInt(m[1]);

  const href    = $('.l-chapter a, .chapter-list a').first().attr('href') || '';
  const firstUrl = href.startsWith('http') ? href : `https://novelbin.com${href}`;

  return { title, author, chapter_count, desc, scraper_args: { start_url: firstUrl || url } };
}

async function extractGeneric(url) {
  const $ = await fetchHtml(url);
  return {
    title: $('h1').first().text().trim() || $('title').text().trim(),
    author: '', chapter_count: 0,
    desc: $('meta[name="description"]').attr('content') || '',
    scraper_args: { url },
  };
}

const EXTRACTORS = { royalroad: extractRoyalroad, novelbin: extractNovelbin };

async function searchTranslations(title) {
  const results = [];
  try {
    const $ = await fetchHtml(`https://www.novelupdates.com/?s=${encodeURIComponent(title)}&post_type=seriesplans`);
    $('.search_main_box_nu a').slice(0, 4).each((_, el) => {
      const href = $(el).attr('href') || '';
      if (href) results.push({ source: 'novelupdates', title: $(el).text().trim(), url: href });
    });
  } catch {}
  try {
    const $ = await fetchHtml(`https://freewebnovel.com/search/?searchkey=${encodeURIComponent(title)}`);
    $('.con-list li a, .search-item a').slice(0, 3).each((_, el) => {
      const href = $(el).attr('href') || '';
      if (href) results.push({ source: 'freewebnovel', title: $(el).text().trim(),
        url: href.startsWith('http') ? href : `https://freewebnovel.com${href}` });
    });
  } catch {}
  return results;
}

async function route(url) {
  const cfg = detectPlatform(url);
  if (!cfg) return { status: 'error', error: `Unknown platform for: ${url}`, url };

  let meta = {};
  try { meta = await (EXTRACTORS[cfg.platform] || extractGeneric)(url); }
  catch (e) { meta = { title: '', author: '', chapter_count: 0, desc: '', scraper_args: { url } }; }

  const result = {
    status: 'ok',
    platform: cfg.platform,
    source_language: cfg.lang,
    needs_translation: cfg.needs_translation,
    needs_browser: cfg.needs_browser,
    needs_login: cfg.needs_login,
    scraper_script: cfg.scraper || '',
    canonical_url: url,
    book_name: meta.title || '',
    author: meta.author || '',
    chapter_count: meta.chapter_count || 0,
    description: meta.desc || '',
    cover_url: meta.cover || '',
    scraper_args: meta.scraper_args || { url },
    translation_candidates: [],
    recommended_action: 'scrape_direct',
  };

  if (cfg.needs_translation && result.book_name) {
    result.translation_candidates = await searchTranslations(result.book_name);
    result.recommended_action = result.translation_candidates.length
      ? 'use_existing_translation' : 'translate_from_source';
  }
  return result;
}

// ── DOCX writer ────────────────────────────────────────────────────────────────

function safeName(name) {
  return (name || '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '').trim() || 'Novel';
}

async function writeDocx(chapters, bookName, fileIdx, outDir, log) {
  const safe    = safeName(bookName);
  const bookDir = path.join(outDir, safe);
  fs.mkdirSync(bookDir, { recursive: true });

  const start    = (fileIdx - 1) * CHAPTERS_PER_FILE + 1;
  const end      = start + CHAPTERS_PER_FILE - 1;
  const filePath = path.join(bookDir, `${safe} ${start}-${end}.docx`);

  const children = [];
  for (const ch of chapters) {
    children.push(new Paragraph({ text: ch.title, heading: HeadingLevel.HEADING_1 }));
    for (const para of ch.paragraphs) {
      children.push(new Paragraph({ children: [new TextRun(para)] }));
    }
  }

  const buf = await Packer.toBuffer(new Document({ sections: [{ children }] }));
  fs.writeFileSync(filePath, buf);
  log(`Written: ${path.basename(filePath)}`);
  return filePath;
}

// ── Google Drive ───────────────────────────────────────────────────────────────

async function uploadToDrive(filePath, folderId, log) {
  if (!APPS_SCRIPT_URL) { log('Drive upload skipped — set APPS_SCRIPT_URL env var first.'); return; }
  try {
    const name = path.basename(filePath);
    const content = fs.readFileSync(filePath).toString('base64');
    const resp = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: name, content, folder_id: folderId }),
    });
    const result = await resp.json();
    if (result.status === 'ok') log(`  Drive: uploaded ${name}`);
    else log(`  Drive upload failed: ${result.message || resp.statusText}`);
  } catch (e) {
    log(`  Drive upload failed: ${e.message}`);
  }
}

// ── Jobs ───────────────────────────────────────────────────────────────────────

const jobs = new Map();

class Job extends EventEmitter {
  constructor() {
    super();
    this.setMaxListeners(20);
    this.stopped = false;
    this.status  = 'running';
  }
  log(msg)  { this.emit('line', msg); }
  done()    { this.status = 'done'; this.emit('line', null); }
  stop()    { this.stopped = true; }
}

// ── Scrapers ───────────────────────────────────────────────────────────────────

async function scrapeNovelbin(startUrl, bookName, startCh, endCh, outDir, job, onBatch) {
  const log = m => job.log(m);
  log(`Output: ${path.join(outDir, safeName(bookName))}`);

  let url      = startUrl;
  let chNum    = startCh;
  let fileIdx  = Math.floor((chNum - 1) / CHAPTERS_PER_FILE) + 1;
  let chapters = [];

  while (url && chNum <= endCh && !job.stopped) {
    log(`Chapter ${chNum}: ${url}`);
    let advanced = false;

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const $ = await fetchHtml(url);
        const title = $('span.chr-text').text().trim() || `Chapter ${chNum}`;
        const paragraphs = [];
        $('#chr-content p').each((_, el) => {
          const t = $(el).text().trim();
          if (t) paragraphs.push(t);
        });

        chapters.push({ title, paragraphs });
        log(`  Saved: ${title} (${paragraphs.length} paragraphs)`);

        if (chapters.length >= CHAPTERS_PER_FILE) {
          const fp = await writeDocx(chapters, bookName, fileIdx, outDir, log);
          if (onBatch) await onBatch(fp);
          chapters = [];
          fileIdx++;
        }

        const next = $('a.js-chapter-nav[data-chapter-nav="next"]');
        const disabled = next.attr('disabled') != null || (next.attr('class') || '').includes('disabled');
        if (!next.length || disabled) { url = null; log('No more chapters.'); break; }
        url = next.attr('data-chapter-url') || next.attr('href') || null;
        chNum++;
        advanced = true;
        await sleep(1000);
        break;
      } catch (e) {
        if (attempt < 3) { log(`  Retry ${attempt}/3: ${e.message}`); await sleep(3000 * attempt); }
        else { log(`  Skipping ch ${chNum}: ${e.message}`); chNum++; advanced = true; }
      }
    }
    if (!advanced) break;
  }

  if (chapters.length > 0) {
    const fp = await writeDocx(chapters, bookName, fileIdx, outDir, log);
    if (onBatch) await onBatch(fp);
  }
  log(`Done. Files in: ${path.join(outDir, safeName(bookName))}`);
}

async function scrapeRoyalroad(startUrl, bookName, startCh, endCh, outDir, job, onBatch) {
  const log = m => job.log(m);
  log(`Output: ${path.join(outDir, safeName(bookName))}`);

  let url      = startUrl;
  let chNum    = startCh;
  let fileIdx  = Math.floor((chNum - 1) / CHAPTERS_PER_FILE) + 1;
  let chapters = [];

  while (url && chNum <= endCh && !job.stopped) {
    log(`Chapter ${chNum}: ${url}`);

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const $ = await fetchHtml(url);

        const title = ($('.chapter-title, h1').first().text().trim()) || `Chapter ${chNum}`;
        const paragraphs = [];
        $('.chapter-inner.chapter-content p, .chapter-content p').each((_, el) => {
          const t = $(el).text().trim();
          if (t) paragraphs.push(t);
        });

        chapters.push({ title, paragraphs });
        log(`  Saved: ${title} (${paragraphs.length} paragraphs)`);

        if (chapters.length >= CHAPTERS_PER_FILE) {
          const fp = await writeDocx(chapters, bookName, fileIdx, outDir, log);
          if (onBatch) await onBatch(fp);
          chapters = [];
          fileIdx++;
        }

        // Next chapter — btn-primary + "next" text + /chapter/ href
        let nextUrl = null;
        $('a').each((_, el) => {
          if (nextUrl) return;
          const cls  = $(el).attr('class') || '';
          const text = $(el).text().trim();
          const href = $(el).attr('href') || '';
          if (cls.includes('btn-primary') && /next/i.test(text) && href.includes('/chapter/'))
            nextUrl = href.startsWith('http') ? href : `https://www.royalroad.com${href}`;
        });
        if (!nextUrl) {
          $('a').each((_, el) => {
            if (nextUrl) return;
            const text = $(el).text().trim();
            const href = $(el).attr('href') || '';
            if (/next\s*chapter/i.test(text) && href.includes('/chapter/'))
              nextUrl = href.startsWith('http') ? href : `https://www.royalroad.com${href}`;
          });
        }

        if (!nextUrl) { log('No more chapters.'); url = null; }
        else { url = nextUrl; chNum++; }
        await sleep(1000);
        break;
      } catch (e) {
        if (attempt < 3) { log(`  Retry ${attempt}/3: ${e.message}`); await sleep(3000 * attempt); }
        else { log(`  Skipping ch ${chNum}: ${e.message}`); chNum++; url = null; }
      }
    }
  }

  if (chapters.length > 0) {
    const fp = await writeDocx(chapters, bookName, fileIdx, outDir, log);
    if (onBatch) await onBatch(fp);
  }
  log(`Done. Files in: ${path.join(outDir, safeName(bookName))}`);
}

async function runJob(jobId, platform, result, startCh, endCh, outDir, useDrive) {
  const job     = jobs.get(jobId);
  const bookName = result.book_name || 'Novel';
  const args     = result.scraper_args || {};

  const onBatch = useDrive
    ? async fp => { job.log(`Uploading to Drive: ${path.basename(fp)}...`); await uploadToDrive(fp, GDRIVE_FOLDER, m => job.log(m)); }
    : null;

  try {
    job.log('='.repeat(50));
    job.log(`Scraping: ${bookName}`);
    job.log(`Platform: ${platform} | Chapters: ${startCh}–${endCh}`);
    job.log(`Output:   ${outDir}`);
    job.log('='.repeat(50));

    if (platform === 'novelbin') {
      await scrapeNovelbin(args.start_url || result.canonical_url, bookName, startCh, endCh, outDir, job, onBatch);
    } else if (platform === 'royalroad') {
      await scrapeRoyalroad(args.url || result.canonical_url, bookName, startCh, endCh, outDir, job, onBatch);
    } else {
      job.log(`\nPlatform '${platform}' uses a browser-based scraper.`);
      job.log(`Script: py Scripts/${result.scraper_script || ''}`);
      job.log('\nSet these values in the script:');
      job.log('-'.repeat(44));
      job.log(`  url           = "${result.canonical_url}"`);
      job.log(`  start_chapter = ${startCh}`);
      job.log(`  end_chapter   = ${endCh}`);
      job.log('-'.repeat(44));
    }
  } catch (e) {
    job.log(`\nError: ${e.message}`);
  } finally {
    job.done();
  }
}

// ── Routes ─────────────────────────────────────────────────────────────────────

app.get('/', (_, res) => res.send(HTML));

app.get('/api/status', (_, res) =>
  res.json({ router: true, http_deps: true, docx: true }));

app.get('/api/drive-status', (_, res) => {
  const ready = !!APPS_SCRIPT_URL;
  res.json({ ready, url_set: ready, folder_id: GDRIVE_FOLDER });
});

app.get('/api/config', (_, res) =>
  res.json({ default_out_dir: DEFAULT_OUT, is_server: IS_SERVER }));

app.post('/api/detect', async (req, res) => {
  const url = (req.body?.url || '').trim();
  if (!url) return res.json({ status: 'error', error: 'No URL provided' });
  try { res.json(await route(url)); }
  catch (e) { res.json({ status: 'error', error: e.message }); }
});

app.post('/api/scrape', (req, res) => {
  const { detection_result: result, start_ch, end_ch, out_dir, upload_to_drive } = req.body;
  if (!result?.platform) return res.status(400).json({ error: 'No detection result' });

  const outDir = path.resolve((out_dir || DEFAULT_OUT).replace(/^~/, HOME));
  fs.mkdirSync(outDir, { recursive: true });

  const id  = randomBytes(4).toString('hex');
  const job = new Job();
  jobs.set(id, job);

  runJob(id, result.platform, result, parseInt(start_ch) || 1, parseInt(end_ch) || 100, outDir, !!upload_to_drive);
  res.json({ job_id: id });
});

app.get('/api/log/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).end();

  res.setHeader('Content-Type',     'text/event-stream');
  res.setHeader('Cache-Control',    'no-cache');
  res.setHeader('X-Accel-Buffering','no');
  res.flushHeaders();

  const onLine = msg => {
    if (msg === null) { res.write('data: "[DONE]"\n\n'); res.end(); }
    else res.write(`data: ${JSON.stringify({ line: msg })}\n\n`);
  };
  job.on('line', onLine);
  req.on('close', () => job.off('line', onLine));
});

app.post('/api/stop/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (job) { job.stop(); res.json({ ok: true }); }
  else res.status(404).json({ ok: false });
});

app.post('/api/open-folder', (req, res) => {
  if (IS_SERVER) return res.json({ ok: false, error: 'Not available on server' });
  const p = path.resolve((req.body?.path || DEFAULT_OUT).replace(/^~/, HOME));
  if (fs.existsSync(p)) { exec(`open "${p}"`); res.json({ ok: true }); }
  else res.json({ ok: false, error: 'Folder not found' });
});

// ── HTML ────────────────────────────────────────────────────────────────────────

const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Novel Scraper</title>
<style>
  :root {
    --bg:      #0f1117; --surface: #1a1d27; --card: #21252f; --border: #2e3244;
    --fg:      #cdd6f4; --muted:   #6c7086; --blue:  #89b4fa; --green: #a6e3a1;
    --red:     #f38ba8; --yellow:  #f9e2af; --teal:  #94e2d5; --r: 8px;
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--fg); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size: 14px; min-height: 100vh; display: flex; flex-direction: column; }

  header { display: flex; align-items: center; gap: 12px; padding: 14px 20px;
    border-bottom: 1px solid var(--border); background: var(--surface); }
  header h1 { font-size: 18px; font-weight: 700; color: var(--blue); }
  .header-status { margin-left: auto; display: flex; gap: 8px; align-items: center; }
  .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--muted); display: inline-block; }
  .dot.ok  { background: var(--green); }
  .dot.err { background: var(--red); }
  .dot-label { font-size: 11px; color: var(--muted); }

  .url-bar { display: flex; gap: 8px; padding: 14px 20px;
    background: var(--surface); border-bottom: 1px solid var(--border); }
  .url-bar input { flex: 1; background: var(--card); border: 1px solid var(--border);
    border-radius: var(--r); color: var(--fg); font-size: 14px; padding: 9px 14px; outline: none; transition: border-color .15s; }
  .url-bar input:focus { border-color: var(--blue); }
  .url-bar input::placeholder { color: var(--muted); }

  button { border: none; border-radius: var(--r); cursor: pointer; font-size: 13px; font-weight: 600;
    padding: 9px 16px; transition: opacity .15s, transform .1s; white-space: nowrap; }
  button:active { transform: scale(.97); }
  button:disabled { opacity: .4; cursor: not-allowed; transform: none; }
  .btn-primary { background: var(--blue);  color: #1e1e2e; }
  .btn-danger  { background: var(--red);   color: #1e1e2e; }
  .btn-ghost   { background: var(--card);  color: var(--fg); border: 1px solid var(--border); }
  .btn-green   { background: var(--green); color: #1e1e2e; }
  .btn-sm { padding: 5px 10px; font-size: 12px; }

  .main { display: grid; grid-template-columns: 1fr 280px; gap: 14px; padding: 14px 20px; flex: 1; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: var(--r); padding: 16px; }
  .card-title { font-size: 11px; font-weight: 700; letter-spacing: .8px; text-transform: uppercase;
    color: var(--muted); margin-bottom: 12px; }
  .info-card { display: flex; flex-direction: column; gap: 10px; }
  .badge { display: inline-flex; gap: 8px; background: var(--surface); border: 1px solid var(--border);
    border-radius: 4px; padding: 4px 10px; font-size: 12px; font-weight: 600; color: var(--blue); align-items: center; align-self: flex-start; }
  .badge .lang-tag { background: var(--blue); color: #1e1e2e; border-radius: 3px; padding: 1px 5px; font-size: 10px; }
  .badge .warn-tag { background: var(--yellow); color: #1e1e2e; border-radius: 3px; padding: 1px 5px; font-size: 10px; }
  .book-title  { font-size: 20px; font-weight: 700; line-height: 1.3; }
  .book-author { color: var(--green); font-size: 14px; }
  .book-meta   { color: var(--muted); font-size: 12px; }
  .book-desc   { color: #bac2de; font-size: 13px; line-height: 1.5; }
  .trans-title { font-size: 11px; color: var(--muted); font-weight: 600; text-transform: uppercase;
    letter-spacing: .6px; margin-bottom: 6px; }
  .trans-item  { display: flex; align-items: center; justify-content: space-between; gap: 10px;
    padding: 7px 10px; background: var(--surface); border: 1px solid var(--border); border-radius: 6px; margin-bottom: 4px; }
  .trans-source { font-size: 10px; color: var(--muted); font-weight: 600; text-transform: uppercase; }
  .trans-name   { font-size: 13px; color: var(--teal); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .config-card { display: flex; flex-direction: column; gap: 12px; }
  .field-label { font-size: 11px; color: var(--muted); font-weight: 600; text-transform: uppercase;
    letter-spacing: .5px; margin-bottom: 4px; }
  .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  input[type=number], .path-input { width: 100%; background: var(--surface); border: 1px solid var(--border);
    border-radius: 6px; color: var(--fg); font-size: 13px; padding: 7px 10px; outline: none; transition: border-color .15s; }
  input[type=number]:focus, .path-input:focus { border-color: var(--blue); }
  .scraper-info { background: var(--surface); border: 1px solid var(--border); border-radius: 6px;
    padding: 8px 10px; font-size: 11px; color: var(--muted); line-height: 1.6; }
  .btn-group { display: flex; flex-direction: column; gap: 6px; }
  .btn-group button { width: 100%; }
  .drive-section { border-top: 1px solid var(--border); padding-top: 12px; }
  .drive-hint { font-size: 11px; color: var(--muted); line-height: 1.6; margin-bottom: 8px; }
  .drive-hint a { color: var(--blue); }
  .drive-hint code { background: var(--surface); padding: 1px 4px; border-radius: 3px; font-family: monospace; }
  .log-section { padding: 0 20px 14px; }
  .log-header  { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
  .log-label   { font-size: 11px; font-weight: 700; letter-spacing: .8px; text-transform: uppercase; color: var(--muted); }
  .log-box { background: #0a0c10; border: 1px solid var(--border); border-radius: var(--r);
    font-family: "Menlo","Fira Code",monospace; font-size: 12px; line-height: 1.6; height: 240px;
    overflow-y: auto; padding: 10px 14px; color: #a6adc8; }
  .log-line { animation: fadeIn .15s ease; }
  .log-line.ok   { color: var(--green); }
  .log-line.err  { color: var(--red); }
  .log-line.info { color: var(--blue); }
  @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
  .progress-wrap { height: 3px; background: var(--border); border-radius: 2px; margin-bottom: 8px;
    overflow: hidden; display: none; }
  .progress-wrap.active { display: block; }
  .progress-bar { height: 100%; background: var(--blue); border-radius: 2px; }
  .progress-bar.indeterminate { width: 30%; animation: slide 1.2s infinite ease-in-out; }
  @keyframes slide { 0% { transform: translateX(-100%); } 100% { transform: translateX(400%); } }
  .spinner { width: 14px; height: 14px; border: 2px solid var(--border); border-top-color: var(--blue);
    border-radius: 50%; animation: spin .7s linear infinite; display: none; }
  .spinner.active { display: inline-block; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .placeholder { color: var(--muted); font-style: italic; }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
</style>
</head>
<body>

<header>
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--blue)" stroke-width="2">
    <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>
  </svg>
  <h1>Novel Scraper</h1>
  <div class="header-status">
    <span class="dot" id="dot-router"></span><span class="dot-label">router</span>
    <span class="dot" id="dot-http"></span><span class="dot-label">http</span>
    <span class="dot" id="dot-docx"></span><span class="dot-label">docx</span>
    <span class="dot" id="dot-drive"></span><span class="dot-label">drive</span>
    <div class="spinner" id="spinner"></div>
  </div>
</header>

<div class="url-bar">
  <input id="url-input" type="text" placeholder="Paste a novel URL (novelbin.com, royalroad.com, webnovel.com, ...)"/>
  <button class="btn-primary" id="detect-btn" onclick="detect()">Detect</button>
</div>

<div class="progress-wrap" id="progress-wrap">
  <div class="progress-bar indeterminate" id="progress-bar"></div>
</div>

<div class="main">
  <div class="card info-card">
    <div class="card-title">Book Info</div>
    <div id="info-placeholder" class="placeholder">Paste a URL above and click Detect.</div>
    <div id="info-content" style="display:none;flex-direction:column;gap:10px">
      <div id="badge" class="badge"></div>
      <div id="book-title" class="book-title"></div>
      <div id="book-author" class="book-author"></div>
      <div id="book-meta" class="book-meta"></div>
      <div id="book-desc" class="book-desc"></div>
      <div id="trans-section" style="display:none">
        <div class="trans-title">English Translation Candidates</div>
        <div id="trans-list"></div>
      </div>
    </div>
  </div>

  <div class="card config-card">
    <div class="card-title">Scrape Config</div>
    <div>
      <div class="field-label">Chapter Range</div>
      <div class="field-row">
        <div><div style="font-size:11px;color:var(--muted);margin-bottom:3px">Start</div>
          <input type="number" id="start-ch" value="1" min="1"/></div>
        <div><div style="font-size:11px;color:var(--muted);margin-bottom:3px">End</div>
          <input type="number" id="end-ch" value="100" min="1"/></div>
      </div>
    </div>
    <div>
      <div class="field-label">Output Directory</div>
      <input type="text" class="path-input" id="out-dir" placeholder="~/Downloads/Novels"/>
    </div>
    <div class="scraper-info" id="scraper-info">Run Detect first.</div>
    <div class="btn-group">
      <button class="btn-green" id="scrape-btn" onclick="startScrape()" disabled>&#9654;&nbsp; Scrape</button>
      <button class="btn-danger" id="stop-btn" onclick="stopScrape()" disabled>&#9632;&nbsp; Stop</button>
      <button class="btn-ghost" onclick="openFolder()">&#128193;&nbsp; Open Output Folder</button>
    </div>
    <div class="drive-section">
      <div class="field-label" style="margin-bottom:8px">Google Drive</div>
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;margin-bottom:8px">
        <input type="checkbox" id="drive-toggle" disabled style="width:15px;height:15px;accent-color:var(--blue)">
        <span style="font-size:13px">Upload after each 100 chapters</span>
      </label>
      <div class="drive-hint" id="drive-hint">
        Folder: <a href="https://drive.google.com/drive/folders/1UyCUOcPTQLGSkII4DoEPd-_gKGZwLO9E" target="_blank">Open in Drive</a>
      </div>
    </div>
  </div>
</div>

<div class="log-section">
  <div class="log-header">
    <span class="log-label">Log</span>
    <button class="btn-ghost btn-sm" onclick="clearLog()">Clear</button>
  </div>
  <div class="log-box" id="log-box"></div>
</div>

<script>
let detectionResult = null, currentJobId = null, eventSource = null;

window.onload = async () => {
  document.getElementById('url-input').addEventListener('keydown', e => { if (e.key === 'Enter') detect(); });

  try {
    const s = await fetch('/api/status').then(r => r.json());
    setDot('dot-router', s.router); setDot('dot-http', s.http_deps); setDot('dot-docx', s.docx);
    log('Ready. Paste a URL above and click Detect.', 'info');
  } catch(e) { log('Cannot reach server: ' + e, 'err'); }

  try {
    const cfg = await fetch('/api/config').then(r => r.json());
    document.getElementById('out-dir').value = cfg.default_out_dir || '~/Downloads/Novels';
  } catch {}

  try {
    const d = await fetch('/api/drive-status').then(r => r.json());
    setDot('dot-drive', d.ready);
    const hint = document.getElementById('drive-hint');
    if (d.service_account || d.oauth_token) {
      document.getElementById('drive-toggle').disabled = false;
      hint.innerHTML = 'Connected. Folder: <a href="https://drive.google.com/drive/folders/1UyCUOcPTQLGSkII4DoEPd-_gKGZwLO9E" target="_blank" style="color:var(--blue)">Open in Drive</a>';
    } else {
      hint.innerHTML = '<b>One-time setup:</b><br>' +
        '1. <a href="https://console.cloud.google.com/iam-admin/serviceaccounts" target="_blank" style="color:var(--blue)">Create a service account</a><br>' +
        '2. Download JSON key → save as <code>service_account.json</code> next to server.js<br>' +
        '3. Share the Drive folder with the service account email';
    }
  } catch {}
};

function setDot(id, ok) {
  const el = document.getElementById(id);
  el.classList.remove('ok','err');
  el.classList.add(ok ? 'ok' : 'err');
}

async function detect() {
  const url = document.getElementById('url-input').value.trim();
  if (!url) { log('Enter a URL first.', 'err'); return; }
  setLoading(true); log(''); log('Detecting: ' + url, 'info');
  try {
    const data = await fetch('/api/detect', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({url}) }).then(r => r.json());
    setLoading(false);
    if (data.status !== 'ok') { log('Error: ' + (data.error || '?'), 'err'); return; }
    detectionResult = data;

    document.getElementById('info-placeholder').style.display = 'none';
    const ic = document.getElementById('info-content');
    ic.style.display = 'flex';

    const badge = document.getElementById('badge');
    const platform = (data.platform || '').toUpperCase();
    const lang = (data.source_language || '').toUpperCase();
    badge.innerHTML = platform + ' <span class="lang-tag">' + lang + '</span>' +
      (data.needs_translation ? ' <span class="warn-tag">needs EN translation</span>' : '');

    document.getElementById('book-title').textContent  = data.book_name || '(title not found)';
    document.getElementById('book-author').textContent = data.author || '';
    document.getElementById('book-meta').textContent   = data.chapter_count ? data.chapter_count.toLocaleString() + ' chapters' : 'chapter count unknown';
    const desc = (data.description || '').slice(0, 280);
    document.getElementById('book-desc').textContent   = desc + (data.description?.length > 280 ? '…' : '');

    if (data.chapter_count > 0) document.getElementById('end-ch').value = data.chapter_count;

    const modeMap = { novelbin: 'Built-in (requests)', royalroad: 'Built-in (requests)', freewebnovel: 'Subprocess (CLI args)' };
    document.getElementById('scraper-info').innerHTML =
      '<div>Script: ' + (data.scraper_script || 'none') + '</div>' +
      '<div>Method: ' + (modeMap[data.platform] || 'Browser script (Chrome required)') + '</div>' +
      (data.needs_login ? '<div>Requires login</div>' : '');

    const candidates = data.translation_candidates || [];
    if (candidates.length) {
      document.getElementById('trans-section').style.display = 'block';
      document.getElementById('trans-list').innerHTML = candidates.slice(0, 6).map(c =>
        '<div class="trans-item"><div class="trans-item-info"><div class="trans-source">' + c.source + '</div>' +
        '<div class="trans-name">' + c.title + '</div></div>' +
        '<button class="btn-ghost btn-sm" onclick="useUrl(' + JSON.stringify(c.url) + ')">Use</button></div>'
      ).join('');
    } else {
      document.getElementById('trans-section').style.display = 'none';
    }

    document.getElementById('scrape-btn').disabled = false;
    log('Platform: ' + platform + ' | Lang: ' + lang + ' | Chapters: ' + (data.chapter_count || '?'), 'ok');
  } catch(e) { setLoading(false); log('Detection error: ' + e, 'err'); }
}

function useUrl(url) { document.getElementById('url-input').value = url; detect(); }

async function startScrape() {
  if (!detectionResult) return;
  const startCh = parseInt(document.getElementById('start-ch').value) || 1;
  const endCh   = parseInt(document.getElementById('end-ch').value)   || 100;
  const outDir  = document.getElementById('out-dir').value.trim();
  if (startCh > endCh) { log('Start must be <= end chapter.', 'err'); return; }

  document.getElementById('scrape-btn').disabled = true;
  document.getElementById('stop-btn').disabled   = false;
  setProgress(true);

  try {
    const res = await fetch('/api/scrape', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ detection_result: detectionResult, start_ch: startCh, end_ch: endCh,
        out_dir: outDir, upload_to_drive: document.getElementById('drive-toggle').checked })
    }).then(r => r.json());

    if (res.error) { log('Error: ' + res.error, 'err'); scrapeDone(); return; }
    currentJobId = res.job_id;
    log('Job started [' + currentJobId + ']', 'info');

    if (eventSource) eventSource.close();
    eventSource = new EventSource('/api/log/' + currentJobId);
    eventSource.onmessage = e => {
      const payload = JSON.parse(e.data);
      if (payload === '[DONE]') { eventSource.close(); scrapeDone(); return; }
      const line = payload.line || '';
      log(line, line.startsWith('Done') || line.includes('Written') ? 'ok' : line.includes('Error') || line.includes('failed') ? 'err' : '');
    };
    eventSource.onerror = () => { eventSource.close(); scrapeDone(); };
  } catch(e) { log('Failed: ' + e, 'err'); scrapeDone(); }
}

async function stopScrape() {
  if (!currentJobId) return;
  document.getElementById('stop-btn').disabled = true;
  await fetch('/api/stop/' + currentJobId, { method:'POST' });
  log('Stop requested...', 'info');
}

function scrapeDone() {
  setProgress(false);
  document.getElementById('scrape-btn').disabled = false;
  document.getElementById('stop-btn').disabled   = true;
  currentJobId = null;
}

async function openFolder() {
  const p = document.getElementById('out-dir').value.trim();
  const r = await fetch('/api/open-folder', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({path: p}) }).then(r => r.json());
  if (!r.ok) log('Could not open folder: ' + (r.error || '?'), 'err');
}

function log(msg, cls) {
  if (!msg && msg !== '') return;
  const box = document.getElementById('log-box');
  if (msg === '') { box.appendChild(document.createElement('br')); box.scrollTop = 9e9; return; }
  const div = document.createElement('div');
  div.className = 'log-line' + (cls ? ' ' + cls : '');
  div.textContent = msg;
  box.appendChild(div);
  box.scrollTop = 9e9;
}
function clearLog()       { document.getElementById('log-box').innerHTML = ''; }
function setLoading(on)   { document.getElementById('spinner').classList.toggle('active', on); document.getElementById('detect-btn').disabled = on; setProgress(on); }
function setProgress(on)  { document.getElementById('progress-wrap').classList.toggle('active', on); }
</script>
</body>
</html>`;

// ── Start ──────────────────────────────────────────────────────────────────────

app.listen(PORT, '0.0.0.0', async () => {
  console.log(`Novel Scraper (JS) running at http://localhost:${PORT}`);
  if (!IS_SERVER) {
    try { const { default: open } = await import('open'); await open(`http://localhost:${PORT}`); } catch {}
  }
});

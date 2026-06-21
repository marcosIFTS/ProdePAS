const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const rootDir = __dirname;
const publicDir = path.join(rootDir, 'public');
const dataDir = path.join(rootDir, 'data');

function send(res, statusCode, body, contentType = 'text/plain; charset=utf-8') {
  res.writeHead(statusCode, {
    'Content-Type': contentType,
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

function sendJson(res, statusCode, data) {
  send(res, statusCode, JSON.stringify(data, null, 2), 'application/json; charset=utf-8');
}

function safeRead(filePath) {
  return fs.promises.readFile(filePath, 'utf8').then((raw) => raw.replace(/^\uFEFF/, ''));
}

async function readTxtJson(fileName) {
  const filePath = path.join(dataDir, fileName);
  const raw = await safeRead(filePath);
  return JSON.parse(raw);
}

function slugify(value) {
  return String(value)
    .normalize('NFD')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');
}

async function getSummary() {
  const general = await readTxtJson('tabla_general.txt');
  const official = await readTxtJson('resultados_oficiales.txt');

  const users = general.rows.map((row) => ({
    name: row.nombre_fantasia || row.user || row.usuario_tab || row.name,
    tabName: row.usuario_tab || row.user || row.tabName || row.slug,
    slug: slugify(row.usuario_tab || row.user || row.slug || row.name),
    points: row.points,
    position: row.position,
  }));

  official.matches = Array.isArray(official.matches)
    ? official.matches.filter((match) => Number(match.number) > 0)
    : [];

  general.rows = Array.isArray(general.rows)
    ? general.rows.filter((row) => Number(row.position) > 0)
    : [];

  return {
    updatedAt: general.updatedAt || official.updatedAt || null,
    general,
    official,
    users,
  };
}

async function getUserDetail(slug) {
  const summary = await getSummary();
  const user = summary.users.find((entry) => entry.slug === slug);
  if (!user) {
    return null;
  }

  const userFile = `${user.slug}.txt`;
  const detail = await readTxtJson(path.join('users', userFile));

  detail.matches = Array.isArray(detail.matches)
    ? detail.matches.filter((match) => Number(match.number) > 0)
    : [];

  return {
    ...user,
    detail,
  };
}

function serveStaticFile(filePath, res) {
  const ext = path.extname(filePath).toLowerCase();
  const contentTypes = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.ico': 'image/x-icon',
  };

  fs.readFile(filePath, (error, data) => {
    if (error) {
      send(res, 404, 'Not found');
      return;
    }

    send(res, 200, data, contentTypes[ext] || 'application/octet-stream');
  });
}

const server = http.createServer(async (req, res) => {
  if (!['GET', 'HEAD'].includes(req.method)) {
    sendJson(res, 405, { error: 'Method not allowed' });
    return;
  }

  const parsedUrl = url.parse(req.url, true);
  const pathname = decodeURIComponent(parsedUrl.pathname || '/');

  if (pathname === '/' || pathname === '/index.html') {
    serveStaticFile(path.join(publicDir, 'index.html'), res);
    return;
  }

  if (pathname.startsWith('/public/')) {
    const filePath = path.join(rootDir, pathname.slice(1));
    if (filePath.startsWith(publicDir)) {
      serveStaticFile(filePath, res);
      return;
    }
  }

  if (pathname === '/api/summary') {
    try {
      const summary = await getSummary();
      sendJson(res, 200, summary);
    } catch (error) {
      sendJson(res, 500, { error: 'No se pudo leer la informacion de TXT', details: error.message });
    }
    return;
  }

  if (pathname === '/api/official') {
    try {
      const official = await readTxtJson('resultados_oficiales.txt');
      sendJson(res, 200, official);
    } catch (error) {
      sendJson(res, 500, { error: 'No se pudo leer los resultados oficiales', details: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/users/')) {
    const slug = pathname.replace('/api/users/', '');
    try {
      const user = await getUserDetail(slug);
      if (!user) {
        sendJson(res, 404, { error: 'Usuario no encontrado' });
        return;
      }
      sendJson(res, 200, user);
    } catch (error) {
      sendJson(res, 500, { error: 'No se pudo leer el detalle del usuario', details: error.message });
    }
    return;
  }

  const assetPath = path.join(publicDir, pathname.replace(/^\//, ''));
  if (assetPath.startsWith(publicDir) && fs.existsSync(assetPath) && fs.statSync(assetPath).isFile()) {
    serveStaticFile(assetPath, res);
    return;
  }

  sendJson(res, 404, { error: 'Ruta no encontrada' });
});

const port = process.env.PORT || 3000;

server.listen(port, () => {
  console.log(`ProdePAS running on http://localhost:${port}`);
});
const state = {
  summary: null,
  activeUser: null,
};

function formatUpdatedAt(value) {
  if (!value) {
    return 'Sin datos';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat('es-AR', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

async function fetchJson(url) {
  const response = await fetch(url, { headers: { Accept: 'application/json' } });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  return response.json();
}

function renderGeneral(rows) {
  const tbody = document.getElementById('generalTable');
  tbody.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>${row.position}</td>
          <td class="fw-semibold">${escapeHtml(row.nombre_fantasia || row.user || row.usuario_tab)}</td>
          <td class="text-end">${row.points}</td>
          <td class="text-end">${row.acertados ?? row.correct ?? 0}</td>
          <td class="text-end">${row.exactos ?? row.exact ?? 0}</td>
          <td class="text-end">${row.played}</td>
        </tr>
      `,
    )
    .join('');
}

function renderChart(rows) {
  const chart = document.getElementById('chart');
  const maxPoints = Math.max(...rows.map((row) => row.points), 1);

  chart.innerHTML = rows
    .map((row) => {
      const width = Math.max((row.points / maxPoints) * 100, 3);
      const displayName = row.nombre_fantasia || row.user || row.usuario_tab;
      return `
        <div class="chart-item">
          <div class="chart-name">${escapeHtml(displayName)}</div>
          <div class="chart-track">
            <div class="chart-bar" style="width: ${width}%"></div>
          </div>
          <div class="chart-points">${row.points}</div>
        </div>
      `;
    })
    .join('');
}

function renderOfficial(matches) {
  const container = document.getElementById('officialList');
  container.innerHTML = matches
    .map(
      (match) => `
        <div class="list-group-item px-0 py-3">
          <div class="d-flex justify-content-between align-items-start gap-3">
            <div>
              <div class="fw-semibold">${escapeHtml(match.local || match.home)} vs ${escapeHtml(match.visitante || match.away)}</div>
              <div class="detail-badge mt-1">${escapeHtml(match.grupo || match.stage)} · ${escapeHtml(match.fecha || match.date)}</div>
            </div>
            <div class="text-end">
              <div class="fs-5 fw-bold">${escapeHtml(match.score || `${match.goles_local ?? ''} - ${match.goles_visitante ?? ''}`)}</div>
              <div class="detail-badge">${escapeHtml(match.status || '')}</div>
            </div>
          </div>
        </div>
      `,
    )
    .join('');
}

function renderUsers(users) {
  const container = document.getElementById('userLinks');
  container.innerHTML = users
    .map(
      (user) => `
        <button class="btn btn-outline-primary user-button" data-user="${escapeHtml(user.slug)}">
          <div class="d-flex justify-content-between align-items-center">
            <span>${escapeHtml(user.name)}${user.tabName && user.tabName !== user.name ? ` <small class="text-secondary">(${escapeHtml(user.tabName)})</small>` : ''}</span>
            <span class="badge text-bg-light">${user.points} pts</span>
          </div>
        </button>
      `,
    )
    .join('');

  container.querySelectorAll('[data-user]').forEach((button) => {
    button.addEventListener('click', () => {
      loadUser(button.getAttribute('data-user'));
    });
  });
}

function renderUserDetail(user) {
  const title = document.getElementById('userDetailTitle');
  const subtitle = document.getElementById('userDetailSubtitle');
  const tbody = document.getElementById('userDetailTable');

  title.textContent = `Detalle de ${user.name}`;
  subtitle.textContent = user.tabName && user.tabName !== user.name
    ? `Usuario ${user.tabName} · Posición ${user.position} · ${user.points} puntos`
    : `Posición ${user.position} · ${user.points} puntos`;

  tbody.innerHTML = user.detail.matches
    .map(
      (match) => `
        <tr>
          <td>${escapeHtml(match.match)}</td>
          <td>${escapeHtml(match.prediction)}</td>
          <td>${escapeHtml(match.official)}</td>
          <td class="text-end fw-semibold">${match.points}</td>
        </tr>
      `,
    )
    .join('');
}

async function loadUser(slug) {
  try {
    const user = await fetchJson(`/api/users/${encodeURIComponent(slug)}`);
    state.activeUser = user;
    renderUserDetail(user);
  } catch (error) {
    const title = document.getElementById('userDetailTitle');
    const subtitle = document.getElementById('userDetailSubtitle');
    title.textContent = 'Detalle del usuario';
    subtitle.textContent = `No se pudo cargar el usuario: ${error.message}`;
  }
}

async function main() {
  state.summary = await fetchJson('/api/summary');
  document.getElementById('updatedAt').textContent = formatUpdatedAt(state.summary.updatedAt);

  renderGeneral(state.summary.general.rows);
  renderChart(state.summary.general.rows);
  renderOfficial(state.summary.official.matches);
  renderUsers(state.summary.users);

  if (state.summary.users.length > 0) {
    await loadUser(state.summary.users[0].slug);
  }
}

main().catch((error) => {
  document.getElementById('updatedAt').textContent = `Error: ${error.message}`;
});
/* ── API client — talks to Laravel backend ─────────────────────────────── */
const API_BASE = 'https://sillar.uk/api';

async function apiFetch(path, options = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    ...options
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`API ${res.status}: ${err}`);
  }
  return res.status === 204 ? null : res.json();
}

const api = {
  // Facilities
  getFacilities:    ()            => apiFetch('/facilities'),
  addFacility:      (data)        => apiFetch('/facilities',        { method:'POST',   body: JSON.stringify(data) }),
  updateFacility:   (id, data)    => apiFetch(`/facilities/${id}`,  { method:'PUT',    body: JSON.stringify(data) }),
  deleteFacility:   (id)          => apiFetch(`/facilities/${id}`,  { method:'DELETE' }),
  importFacilities: (list)        => apiFetch('/facilities/import', { method:'POST',   body: JSON.stringify({ facilities: list }) }),

  // Immunization Data
  getData:  (facId, year)        => apiFetch(`/data/${facId}/${year}`),
  saveData: (facId, year, data)  => apiFetch(`/data/${facId}/${year}`, { method:'POST', body: JSON.stringify(data) }),
  getAllData:(facId)              => apiFetch(`/data/${facId}`),

  // Stock Log
  getStock:  (facId, year, month) => apiFetch(`/stock/${facId}/${year}/${month}`),
  saveStock: (facId, year, month, data) => apiFetch(`/stock/${facId}/${year}/${month}`, { method:'POST', body: JSON.stringify(data) }),
  getStockYear: (facId, year)     => apiFetch(`/stock/${facId}/${year}`),
};

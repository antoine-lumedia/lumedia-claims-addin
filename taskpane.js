/**
 * Claims Migration task-pane add-in — Supabase + Graph version
 *
 * Architecture (since 2026-05-24):
 *   The add-in talks directly to the workbook via the Excel JS API for
 *   reading the active sheet's pending migrations and clearing dropdowns,
 *   and POSTs the batch of migrations to a single Supabase Edge Function
 *   endpoint (`migrate-rows`). No Office Scripts, no Power Automate.
 *
 * Config: only one named cell on the hidden _FlowConfig sheet matters now:
 *   migrate_rows_url — the Supabase function URL,
 *                      e.g. https://<project>.supabase.co/functions/v1/migrate-rows
 *   migrate_rows_anon_key — the Supabase anon (or service-role) bearer used
 *                      for the call. Anon is fine; the function trusts the
 *                      request because it's coming from inside the tracker.
 *
 * Flow:
 *   idle  → user clicks "Check pending changes"
 *   pending → list rendered, user fills any required block_reason fields,
 *             clicks Apply (enabled only when all required reasons are valid)
 *   running → one batched POST to migrate-rows; live progress on results
 *   result  → success summary, or partial-failure state
 *
 * Pending-migration detection lives entirely client-side:
 *   - Find the active sheet (must be one of the pipeline sheets)
 *   - Find its table (tblSOAReceived, tblOTReceived, etc.)
 *   - Walk rows looking for non-empty status_action cells
 *   - Map status_action → target sheet (e.g. SOA_Received + "Send to brand" → SOA_Claimed)
 *   - Group by contract_no
 *
 * Author: Claude (Anthropic), originally based on the Power Automate version.
 */

'use strict';

// ---------- Status-action → target-sheet mapping ----------
// Must match the dropdown values defined in the workbook's data validation.
// Convention: "→ Xxx" = forward migration to Xxx; "← Xxx" = back to Xxx.
const ACTION_MAP = {
  'SOA_Received': {
    '→ Claimed': 'SOA_Claimed',
    '→ Stuck': 'SOA_Stuck',
  },
  'SOA_Claimed': {
    '→ Creditable': 'SOA_Creditable',
    '→ Stuck': 'SOA_Stuck',
    '← Received': 'SOA_Received',
  },
  'SOA_Creditable': {
    '→ Done': 'SOA_Done',
    '→ Stuck': 'SOA_Stuck',
    '← Claimed': 'SOA_Claimed',
  },
  'SOA_Stuck': {
    '→ Received': 'SOA_Received',
    '→ Claimed': 'SOA_Claimed',
    '→ Creditable': 'SOA_Creditable',
    '→ Done': 'SOA_Done',
  },
  'OT_Received': {
    '→ Claimed': 'OT_Claimed',
    '→ Stuck': 'OT_Stuck',
  },
  'OT_Claimed': {
    '→ Creditable': 'OT_Creditable',
    '→ Stuck': 'OT_Stuck',
    '← Received': 'OT_Received',
  },
  'OT_Creditable': {
    '→ Done': 'OT_Done',
    '→ Stuck': 'OT_Stuck',
    '← Claimed': 'OT_Claimed',
  },
  'OT_Stuck': {
    '→ Received': 'OT_Received',
    '→ Claimed': 'OT_Claimed',
    '→ Creditable': 'OT_Creditable',
    '→ Done': 'OT_Done',
  },
};

const TABLE_NAME_FOR = {
  'SOA_Received': 'tblSOAReceived',
  'SOA_Claimed': 'tblSOAClaimed',
  'SOA_Creditable': 'tblSOACreditable',
  'SOA_Done': 'tblSOADone',
  'SOA_Stuck': 'tblSOAStuck',
  'OT_Received': 'tblOTReceived',
  'OT_Claimed': 'tblOTClaimed',
  'OT_Creditable': 'tblOTCreditable',
  'OT_Done': 'tblOTDone',
  'OT_Stuck': 'tblOTStuck',
};

let pendingState = null;
let userInputs = {};

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    showError('This add-in only runs in Excel.');
    return;
  }
  document.getElementById('refresh-btn').onclick = checkPending;
  document.getElementById('apply-btn').onclick = applyMigrations;
  document.getElementById('cancel-btn').onclick = () => showState('idle');
  document.getElementById('done-btn').onclick = () => showState('idle');
  document.getElementById('error-retry-btn').onclick = applyMigrations;
  document.getElementById('error-back-btn').onclick = () => showState('idle');
});

function showState(name) {
  for (const s of ['idle', 'pending', 'running', 'result', 'error']) {
    document.getElementById('state-' + s).hidden = (s !== name);
  }
}
function showError(msg) {
  document.getElementById('error-content').textContent = msg;
  showState('error');
}

async function readNamedCell(name) {
  return Excel.run(async ctx => {
    const range = ctx.workbook.names.getItem(name).getRange();
    range.load('values');
    await ctx.sync();
    return range.values[0][0];
  });
}

async function getCurrentUserEmail() {
  // Excel context doesn't reliably expose user email. Returning empty string
  // means the audit log will record actor='unknown'. Phase 2: prompt user
  // once and cache in a named cell.
  return '';
}

// ---------- step 1: scan pending migrations (client-side, no server call) ----------
async function checkPending() {
  showState('idle');
  document.getElementById('refresh-btn').disabled = true;
  try {
    pendingState = await scanPendingMigrations();
    if (!pendingState.pending || pendingState.pending.length === 0) {
      document.getElementById('result-content').innerHTML =
        '<strong>Nothing to apply.</strong> Set the <code>status_action</code> dropdown on a row first.';
      showState('result');
      return;
    }
    renderPendingUi();
    showState('pending');
  } catch (e) {
    showError(e.message || String(e));
  } finally {
    document.getElementById('refresh-btn').disabled = false;
  }
}

async function scanPendingMigrations() {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    await ctx.sync();

    const sheetName = sheet.name;
    const tableName = TABLE_NAME_FOR[sheetName];
    if (!tableName || !ACTION_MAP[sheetName]) {
      throw new Error(`Active sheet "${sheetName}" is not a pipeline sheet. Switch to one of: ${Object.keys(ACTION_MAP).join(', ')}.`);
    }

    const table = sheet.tables.getItem(tableName);
    const range = table.getDataBodyRange();
    range.load(['values']);
    await ctx.sync();

    // Column positions per the new tracker schema (2026-05-25 rebuild):
    // Both SOA and OT pipeline sheets share these positions for the columns
    // the add-in scans:
    //   A (0) status_action   — dropdown
    //   B (1) row_key
    //   C (2) block_reason
    //   D (3) customer
    //   E (4) contract_no
    //   F (5) message_id
    //   G (6) action_name
    const STATUS_COL = 0;
    const BLOCK_REASON_COL = 2;
    const CONTRACT_COL = 4;
    const ACTION_COL = 6;

    const values = range.values || [];
    const pipeline = sheetName.startsWith('SOA_') ? 'SOA' : 'OT';

    const grouped = new Map();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const action = String(row[STATUS_COL] ?? '').trim();
      if (!action) continue;

      const target = ACTION_MAP[sheetName][action];
      if (!target) {
        throw new Error(`Row ${i + 2}: unknown status_action "${action}" on sheet ${sheetName}`);
      }
      const contractNo = String(row[CONTRACT_COL] ?? '').trim();
      if (!contractNo) {
        throw new Error(`Row ${i + 2}: status_action set but contract_no is empty`);
      }
      const actionName = String(row[ACTION_COL] ?? '').trim();
      const existingBlockReason = String(row[BLOCK_REASON_COL] ?? '').trim();

      if (grouped.has(contractNo)) {
        const existing = grouped.get(contractNo);
        if (existing.status_action !== action) {
          throw new Error(
            `Contract ${contractNo} has rows with different status_action values. ` +
            `Found "${existing.status_action}" and "${action}". Set them all to the same value.`,
          );
        }
      } else {
        grouped.set(contractNo, {
          contract_no: contractNo,
          action: actionName,
          status_action: action,
          target_sheet: target,
          existing_block_reason: existingBlockReason,
        });
      }
    }

    return {
      sheet_name: sheetName,
      pipeline,
      pending: Array.from(grouped.values()),
    };
  });
}

// ---------- step 2: render confirmation list ----------
function renderPendingUi() {
  document.getElementById('active-sheet').textContent = pendingState.sheet_name;
  document.getElementById('active-pipeline').textContent = pendingState.pipeline;
  document.getElementById('global-comment').value = '';
  userInputs = {};

  const list = document.getElementById('pending-list');
  list.innerHTML = '';

  pendingState.pending.forEach(item => {
    userInputs[item.contract_no] = {
      block_reason: item.existing_block_reason || '',
      comment: '',
    };

    const div = document.createElement('div');
    div.className = 'pending-item';
    const isStuckTarget = item.target_sheet.endsWith('_Stuck');
    if (isStuckTarget) div.classList.add('stuck-target');

    const header = document.createElement('div');
    header.className = 'row-header';
    header.textContent = `Contract ${item.contract_no}: ${item.action}`;
    div.appendChild(header);

    const meta = document.createElement('div');
    meta.className = 'row-meta';
    meta.textContent = `${pendingState.sheet_name} → ${item.target_sheet}`;
    div.appendChild(meta);

    if (isStuckTarget) {
      const wrap = document.createElement('div');
      wrap.className = 'block-reason-wrap';
      const label = document.createElement('label');
      label.textContent = 'Block reason';
      label.className = 'required';
      label.htmlFor = `br-${item.contract_no}`;
      const ta = document.createElement('textarea');
      ta.id = `br-${item.contract_no}`;
      ta.rows = 2;
      ta.placeholder = 'Why is this blocked? (min 3 chars)';
      ta.value = item.existing_block_reason || '';
      const msg = document.createElement('div');
      msg.className = 'validation-msg';
      ta.oninput = () => {
        userInputs[item.contract_no].block_reason = ta.value;
        const valid = ta.value.trim().length >= 3;
        ta.classList.toggle('invalid', !valid);
        msg.textContent = valid ? '' : 'Required — at least 3 characters.';
        updateApplyButton();
      };
      ta.dispatchEvent(new Event('input'));
      wrap.appendChild(label);
      wrap.appendChild(ta);
      wrap.appendChild(msg);
      div.appendChild(wrap);
    }
    list.appendChild(div);
  });

  updateApplyButton();
}

function updateApplyButton() {
  const allValid = pendingState.pending.every(item => {
    if (!item.target_sheet.endsWith('_Stuck')) return true;
    const r = (userInputs[item.contract_no].block_reason || '').trim();
    return r.length >= 3;
  });
  document.getElementById('apply-btn').disabled = !allValid;
}

// ---------- step 3: apply (single batched POST) ----------
async function applyMigrations() {
  if (!pendingState) {
    showError('No pending state. Click "Check pending changes" first.');
    return;
  }
  showState('running');
  const progress = document.getElementById('progress');
  progress.innerHTML = '';

  const globalComment = document.getElementById('global-comment').value.trim();
  let migrateUrl, anonKey;
  try {
    migrateUrl = await readNamedCell('migrate_rows_url');
    anonKey = await readNamedCell('migrate_rows_anon_key');
    if (!migrateUrl || migrateUrl.toString().startsWith('<<')) {
      throw new Error('migrate_rows_url not configured on the _FlowConfig sheet.');
    }
    if (!anonKey || anonKey.toString().startsWith('<<')) {
      throw new Error('migrate_rows_anon_key not configured on the _FlowConfig sheet.');
    }
  } catch (e) {
    showError(e.message || String(e));
    return;
  }

  const actor = await getCurrentUserEmail();
  const migrations = pendingState.pending.map(item => {
    const isStuckTarget = item.target_sheet.endsWith('_Stuck');
    return {
      contract_no: item.contract_no,
      pipeline: pendingState.pipeline,
      from_sheet: pendingState.sheet_name,
      to_sheet: item.target_sheet,
      comment: isStuckTarget ? '' : globalComment,
      block_reason: isStuckTarget ? userInputs[item.contract_no].block_reason.trim() : null,
    };
  });

  const liByContract = {};
  for (const m of migrations) {
    const li = document.createElement('li');
    li.textContent = `Contract ${m.contract_no} → ${m.to_sheet}…`;
    progress.appendChild(li);
    liByContract[m.contract_no] = li;
  }

  let results;
  try {
    const res = await fetch(migrateUrl.toString(), {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + anonKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ actor, migrations }),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Function returned ${res.status}: ${text}`);
    }
    const body = await res.json();
    results = body.results || [];
  } catch (e) {
    for (const m of migrations) {
      const li = liByContract[m.contract_no];
      li.classList.add('err');
      li.textContent = `Contract ${m.contract_no}: ${e.message || e}`;
    }
    document.getElementById('result-content').innerHTML =
      `<strong>Network/server error.</strong> No migrations were applied.<br>${e.message || e}`;
    showState('result');
    return;
  }

  const succeeded = [];
  const failed = [];
  for (const r of results) {
    const li = liByContract[r.contract_no];
    if (!li) continue;
    if (r.status === 'ok') {
      const target = pendingState.pending.find(p => p.contract_no === r.contract_no)?.target_sheet;
      li.classList.add('ok');
      li.textContent = `Contract ${r.contract_no} → ${target}: done (${r.rows} rows)`;
      succeeded.push(r.contract_no);
    } else {
      li.classList.add('err');
      li.textContent = `Contract ${r.contract_no}: ${r.error}`;
      failed.push(r);
    }
  }

  if (succeeded.length > 0) {
    try {
      await clearStatusActions(pendingState.sheet_name, succeeded);
    } catch (e) {
      const li = document.createElement('li');
      li.classList.add('err');
      li.textContent = `Warning: could not clear dropdowns (${e.message}). Clear manually.`;
      progress.appendChild(li);
    }
  }

  const total = pendingState.pending.length;
  let html;
  if (failed.length === 0) {
    html = `<strong>${total} migration${total === 1 ? '' : 's'} applied successfully.</strong>`;
  } else if (succeeded.length === 0) {
    html = `<strong>All ${total} migrations failed.</strong><br>See list above. Try again or check Supabase logs.`;
  } else {
    html = `<strong>${succeeded.length} of ${total} succeeded.</strong><br>`
         + `${failed.length} failed — see list above. Dropdowns for failed rows are unchanged so you can retry.`;
  }
  document.getElementById('result-content').innerHTML = html;
  showState('result');
}

// ---------- helper: clear status_action cells client-side ----------
async function clearStatusActions(sheetName, contractNos) {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const tableName = TABLE_NAME_FOR[sheetName];
    const table = sheet.tables.getItem(tableName);
    const range = table.getDataBodyRange();
    range.load(['values']);
    await ctx.sync();

    const CONTRACT_COL = 4;
    const STATUS_COL = 0;
    const wanted = new Set(contractNos.map(c => String(c).trim()));

    for (let i = 0; i < range.values.length; i++) {
      const contract = String(range.values[i][CONTRACT_COL] ?? '').trim();
      if (wanted.has(contract)) {
        const cell = range.getCell(i, STATUS_COL);
        cell.values = [['']];
      }
    }
    await ctx.sync();
  });
}

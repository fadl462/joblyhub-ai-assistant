import * as XLSX from "xlsx";

export interface Env {
  DB: D1Database;
  ASSETS: Fetcher;
  ADMIN_TOKEN: string;
  RESEND_API_KEY: string;
  FROM_EMAIL: string;
  FROM_NAME: string;
  TEST_EMAIL: string;
  APP_BASE_URL: string;
  DAILY_MAX: string;
  SEND_START_HOUR: string;
  SEND_END_HOUR: string;
  SEND_END_MINUTE: string;
}

interface Organization {
  id: number; row_number: number | null; day_batch: string | null; day_number: number | null;
  batch_position: number | null; planned_send_date: string | null; planned_followup_date: string | null;
  name: string; sector: string | null; email: string; is_ngo: number; consent_status: string;
  consent_record: string | null; approved: number; opted_out: number; unsubscribe_token: string;
  email_sent: number; date_sent: string | null; status: string; response_received: number;
  response_sentiment: string | null; responded_at: string | null; followup_sent: number;
  followup_sent_at: string | null; account_created: number; job_posted: number; became_paid: number;
  notes: string | null; created_at: string; updated_at: string;
}


// Safe, repeatable bootstrap: creates missing D1 tables/indexes on first API or scheduled request.
// This avoids a blank-database failure if the D1 migration was not applied in the dashboard.
const BOOTSTRAP_SQL = String.raw`
CREATE TABLE IF NOT EXISTS organizations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  row_number INTEGER,
  day_batch TEXT,
  day_number INTEGER,
  batch_position INTEGER,
  planned_send_date TEXT,
  planned_followup_date TEXT,
  name TEXT NOT NULL,
  sector TEXT,
  email TEXT NOT NULL,
  email_normalized TEXT NOT NULL UNIQUE,
  is_ngo INTEGER NOT NULL DEFAULT 0,
  consent_status TEXT NOT NULL DEFAULT 'unknown',
  consent_record TEXT,
  approved INTEGER NOT NULL DEFAULT 0,
  opted_out INTEGER NOT NULL DEFAULT 0,
  unsubscribe_token TEXT NOT NULL UNIQUE,
  email_sent INTEGER NOT NULL DEFAULT 0,
  date_sent TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  response_received INTEGER NOT NULL DEFAULT 0,
  response_sentiment TEXT,
  responded_at TEXT,
  followup_sent INTEGER NOT NULL DEFAULT 0,
  followup_sent_at TEXT,
  account_created INTEGER NOT NULL DEFAULT 0,
  job_posted INTEGER NOT NULL DEFAULT 0,
  became_paid INTEGER NOT NULL DEFAULT 0,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_org_status ON organizations(status);
CREATE INDEX IF NOT EXISTS idx_org_planned_send ON organizations(planned_send_date);
CREATE INDEX IF NOT EXISTS idx_org_followup ON organizations(planned_followup_date);
CREATE INDEX IF NOT EXISTS idx_org_sector ON organizations(sector);

CREATE TABLE IF NOT EXISTS send_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  organization_id INTEGER,
  kind TEXT NOT NULL DEFAULT 'initial',
  mode TEXT NOT NULL DEFAULT 'test',
  sent_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  to_address TEXT,
  success INTEGER NOT NULL DEFAULT 0,
  provider_message_id TEXT,
  error_message TEXT,
  idempotency_key TEXT NOT NULL UNIQUE,
  FOREIGN KEY (organization_id) REFERENCES organizations(id)
);
CREATE INDEX IF NOT EXISTS idx_send_logs_time ON send_logs(sent_at);

CREATE TABLE IF NOT EXISTS settings (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  is_active INTEGER NOT NULL DEFAULT 0,
  test_mode INTEGER NOT NULL DEFAULT 1,
  test_email TEXT NOT NULL DEFAULT 'fadl_rahman3@yahoo.com',
  from_name TEXT NOT NULL DEFAULT 'JoblyHub',
  from_email TEXT NOT NULL DEFAULT 'business@joblyhub.com',
  daily_min INTEGER NOT NULL DEFAULT 10,
  daily_max INTEGER NOT NULL DEFAULT 15,
  window_start_hour INTEGER NOT NULL DEFAULT 9,
  window_end_hour INTEGER NOT NULL DEFAULT 16,
  window_end_minute INTEGER NOT NULL DEFAULT 30,
  followups_enabled INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
INSERT OR IGNORE INTO settings (id) VALUES (1);

CREATE TABLE IF NOT EXISTS daily_run_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  run_date TEXT NOT NULL UNIQUE,
  scheduled_count INTEGER NOT NULL DEFAULT 0,
  sent_count INTEGER NOT NULL DEFAULT 0,
  completed INTEGER NOT NULL DEFAULT 0,
  started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  completed_at TEXT
);

CREATE TABLE IF NOT EXISTS notifications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  message TEXT NOT NULL,
  level TEXT NOT NULL DEFAULT 'info',
  is_read INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_notifications_created ON notifications(created_at);`;
let schemaReady: Promise<unknown> | null = null;
async function ensureSchema(env: Env) {
  if (!schemaReady) schemaReady = env.DB.exec(BOOTSTRAP_SQL).catch((error) => { schemaReady = null; throw error; });
  await schemaReady;
}

const json = (data: unknown, status = 200) => new Response(JSON.stringify(data), {
  status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }
});
const html = (body: string, status = 200) => new Response(body, { status, headers: { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" } });
const utcNow = () => new Date().toISOString();
const dateKey = (d = new Date()) => d.toISOString().slice(0, 10); // Ghana is UTC+0
const bool = (v: unknown) => v === true || v === 1 || v === "1" || String(v).toLowerCase() === "true" || String(v).toLowerCase() === "yes" || String(v).toLowerCase() === "y";
const safe = (v: unknown, max = 500) => String(v ?? "").trim().slice(0, max);
const emailValid = (s: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s) && s.length <= 254;

function csvRows(input: string): string[][] {
  const rows: string[][] = []; let row: string[] = []; let cell = ""; let quoted = false;
  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (ch === '"' && quoted && input[i + 1] === '"') { cell += '"'; i++; }
    else if (ch === '"') quoted = !quoted;
    else if (ch === "," && !quoted) { row.push(cell); cell = ""; }
    else if ((ch === "\n" || ch === "\r") && !quoted) {
      if (ch === "\r" && input[i + 1] === "\n") i++;
      row.push(cell); if (row.some(c => c.trim())) rows.push(row); row = []; cell = "";
    } else cell += ch;
  }
  row.push(cell); if (row.some(c => c.trim())) rows.push(row);
  return rows;
}
function headerKey(s: string) { const raw = s.trim().toLowerCase(); if (raw === "#") return "row_number"; return raw.replace(/[\s#./()-]+/g, "_").replace(/[^a-z0-9_]/g, "").replace(/^_+|_+$/g, ""); }
function normalizeDate(value: string) { if (!value) return null; const s=value.trim(); if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; let m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/); if (m) { const a=Number(m[1]), b=Number(m[2]), y=Number(m[3]); const day=a>12?a:b>12?b:a; const month=a>12?b:b>12?a:b; if(month>=1&&month<=12&&day>=1&&day<=31) return `${y}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`; } const parsed=new Date(s); return Number.isNaN(parsed.getTime())?null:parsed.toISOString().slice(0,10); }
function pick(row: Record<string,string>, ...keys: string[]) { for (const k of keys) if (row[k] !== undefined && row[k] !== "") return row[k]; return ""; }
function parseCsv(text: string) {
  const rows = csvRows(text.replace(/^\uFEFF/, ""));
  if (rows.length < 2) throw new Error("CSV must include a header row and at least one data row.");
  const headers = rows[0].map(headerKey);
  return rows.slice(1).map((values, i) => {
    const r: Record<string,string> = {}; headers.forEach((h, j) => r[h] = String(values[j] ?? "").trim());
    const email = pick(r, "email", "email_address", "contact_email").toLowerCase();
    const name = pick(r, "name", "organization", "organisation", "company", "organization_name", "organisation_name");
    const consentRaw = pick(r, "consent_status", "consent", "marketing_consent", "written_consent").toLowerCase();
    const consent = ["yes","true","1","written","consented"].includes(consentRaw) ? "written" : consentRaw === "no" || consentRaw === "false" || consentRaw === "0" || consentRaw === "denied" ? "denied" : "unknown";
    const dayBatch = pick(r, "day_batch", "day", "batch");
    const batchMatch = dayBatch.match(/Day\s+(\d+)\s*[·.]\s*#(\d+)/i);
    return {
      row_number: Number(pick(r, "row_number", "row", "number", "#")) || i + 1,
      day_batch: dayBatch,
      day_number: Number(pick(r, "day_number")) || (batchMatch ? Number(batchMatch[1]) : null),
      batch_position: Number(pick(r, "batch_position", "position")) || (batchMatch ? Number(batchMatch[2]) : null),
      planned_send_date: normalizeDate(pick(r, "planned_send_date", "send_date") || ""),
      planned_followup_date: normalizeDate(pick(r, "planned_followup_date", "followup_date", "follow_up_date") || ""),
      name, sector: pick(r, "sector", "industry", "category") || null, email,
      is_ngo: bool(pick(r, "is_ngo", "ngo")) ? 1 : 0,
      consent_status: consent,
      consent_record: pick(r, "consent_record", "consent_evidence", "consent_date", "consent_source") || null,
      approved: bool(pick(r, "approved", "approved_for_sending", "send_approved")) ? 1 : 0,
      opted_out: bool(pick(r, "opted_out", "unsubscribed", "do_not_contact")) || /opt.?out|unsubscribed|do not contact/i.test(pick(r,"status")) ? 1 : 0,
      email_sent: bool(pick(r, "email_sent", "sent", "email_sent_")) || /^(sent|bounced|out of office|replied|mailbox full|no response)$/i.test(pick(r,"status")) ? 1 : 0,
      date_sent: normalizeDate(pick(r, "date_sent", "sent_date") || "") ? `${normalizeDate(pick(r, "date_sent", "sent_date") || "")}T00:00:00.000Z` : null,
      status: pick(r, "status") || (bool(pick(r,"email_sent","sent")) ? "Sent" : "Pending"),
      response_received: bool(pick(r, "response_received", "replied", "response")) ? 1 : 0,
      response_sentiment: pick(r, "response_sentiment", "sentiment") || null,
      responded_at: pick(r, "responded_at", "response_date") || null,
      followup_sent: bool(pick(r, "followup_sent", "follow_up_sent")) ? 1 : 0,
      followup_sent_at: pick(r, "followup_sent_at", "follow_up_sent_at") || null,
      account_created: bool(pick(r, "account_created")) ? 1 : 0,
      job_posted: bool(pick(r, "job_posted")) ? 1 : 0,
      became_paid: bool(pick(r, "became_paid", "paid_customer")) ? 1 : 0,
      notes: pick(r, "notes", "note") || null,
    };
  });
}

function parseXlsx(buffer: ArrayBuffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true, raw: false });
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
    const headerIndex = rows.findIndex((row) => {
      const keys = row.map((v) => headerKey(String(v ?? "")));
      return keys.some((k) => ["email", "email_address", "contact_email"].includes(k)) &&
        keys.some((k) => ["name", "organization", "organisation", "company", "organization_name", "organisation_name"].includes(k));
    });
    if (headerIndex < 0 || rows.length <= headerIndex + 1) continue;
    const headers = rows[headerIndex].map((v) => headerKey(String(v ?? "")));
    const orgCol = headers.findIndex((k) => ["name", "organization", "organisation", "company", "organization_name", "organisation_name"].includes(k));
    const ngoCol = headers.findIndex((k) => ["is_ngo", "ngo"].includes(k));
    // Preserve the original workbook's highlighted NGO marker where its fill is available.
    const csv = rows.slice(headerIndex).map((row, offset) => {
      const values = row.map((v) => String(v ?? ""));
      if (offset === 0 && ngoCol < 0) values.push("is_ngo");
      if (offset > 0 && ngoCol < 0 && orgCol >= 0) {
        const cell = sheet[XLSX.utils.encode_cell({ r: headerIndex + offset, c: orgCol })] as any;
        const rgb = String(cell?.s?.fgColor?.rgb ?? cell?.s?.fill?.fgColor?.rgb ?? "").toUpperCase();
        values.push(rgb.endsWith("FFF8E7") ? "yes" : "no");
      }
      return values.map((v) => `"${v.replace(/"/g, '""')}"`).join(",");
    }).join("\n");
    const parsed = parseCsv(csv);
    if (ngoCol < 0) {
      // parseCsv only sees is_ngo when the synthetic column is appended above.
      // The legacy tracker is often color-coded; XLSX readers that omit styles default safely to false.
    }
    return parsed;
  }
  throw new Error("Could not find a worksheet with organization/name and email columns. Export a tracker sheet with a header row and try again.");
}

async function requireAdmin(req: Request, env: Env) {
  const provided = req.headers.get("x-admin-token") || "";
  if (!env.ADMIN_TOKEN || provided.length !== env.ADMIN_TOKEN.length) return false;
  let diff = 0; for (let i = 0; i < provided.length; i++) diff |= provided.charCodeAt(i) ^ env.ADMIN_TOKEN.charCodeAt(i);
  return diff === 0;
}
function adminError() { return json({ error: "Unauthorized. Enter your admin token in the dashboard." }, 401); }
async function notify(env: Env, title: string, message: string, level = "info") {
  await env.DB.prepare("INSERT INTO notifications(title,message,level) VALUES(?,?,?)").bind(title, message.slice(0, 1500), level).run();
}
async function sendProviderEmail(env: Env, to: string, subject: string, body: string, idempotencyKey: string, replyTo?: string, fromEmail?: string, fromName?: string) {
  if (!env.RESEND_API_KEY) throw new Error("RESEND_API_KEY is not configured.");
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { "Authorization": `Bearer ${env.RESEND_API_KEY}`, "Content-Type": "application/json", "Idempotency-Key": idempotencyKey },
    body: JSON.stringify({ from: `${fromName || env.FROM_NAME || "JoblyHub"} <${fromEmail || env.FROM_EMAIL}>`, to: [to], subject, html: body, ...(replyTo ? { reply_to: replyTo } : {}) })
  });
  const data: any = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(`Email provider error ${response.status}: ${safe(data.message || data.name || "Request failed", 300)}`);
  return String(data.id || "");
}
function escapeHtml(s: string) { return s.replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]!)); }
function emailBody(org: Organization, kind: "initial" | "followup", unsubscribeUrl: string) {
  const name = escapeHtml(org.name || "Team");
  const intro = kind === "initial" ? `We are pleased to introduce <a href="https://joblyhub.com" style="color:#00A69C">JoblyHub</a>, a Ghanaian employment and recruitment platform built to help organizations connect with qualified talent through transparency, technology, and recruitment quality assurance.` : `Just following up on our earlier note in case it slipped past you. We would still love to welcome your organization to the JoblyHub employer community.`;
  return `<table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;font-family:Arial,Helvetica,sans-serif;color:#1f2d2d"><tr><td style="background:#f4f9f9;padding:26px;text-align:center"><div style="font-size:24px;font-weight:bold;color:#006666">➤ JoblyHub</div><div style="margin-top:8px;color:#00A69C;font-weight:bold">CONNECTING OPPORTUNITIES. BUILDING TRUST.</div><div style="margin-top:5px;font-size:13px;color:#4d5e5e">Reach Qualified Talent. Hire with Confidence.</div></td></tr><tr><td style="padding:24px;font-size:14px;line-height:1.7"><p>Greetings from JoblyHub,</p><p>Dear ${name} team,</p><p>${intro}</p><p>JoblyHub helps employers strengthen their visibility and connect with job seekers across Ghana and beyond. All vacancies are reviewed before publication.</p><p><a href="https://joblyhub.com/register" style="display:inline-block;background:#00A69C;color:#fff;text-decoration:none;padding:12px 18px;border-radius:5px;font-weight:bold">CREATE EMPLOYER ACCOUNT</a></p><p>We would be delighted to welcome your organization to the growing JoblyHub employer community. Please reply to this email if you have any questions.</p><p>Kind regards,<br><strong>Business Development Team</strong><br>JoblyHub<br><a href="mailto:${envSafeEmail("business@joblyhub.com")}">business@joblyhub.com</a><br>+233 20 411 6245</p><hr style="border:0;border-top:1px solid #e3e9ea"><p style="font-size:11px;color:#667">You are receiving this message because your organization was included in JoblyHub's approved business outreach list. <a href="${unsubscribeUrl}">Unsubscribe / do not contact me</a>.</p></td></tr></table>`;
}
function envSafeEmail(s: string) { return s.replace(/[<>"']/g, ""); }

async function logSend(env: Env, orgId: number, kind: string, mode: string, to: string, success: boolean, key: string, providerId = "", error = "") {
  await env.DB.prepare(`INSERT INTO send_logs(organization_id,kind,mode,to_address,success,provider_message_id,error_message,idempotency_key) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(idempotency_key) DO UPDATE SET success=excluded.success,provider_message_id=COALESCE(excluded.provider_message_id,send_logs.provider_message_id),error_message=excluded.error_message,sent_at=CURRENT_TIMESTAMP`)
    .bind(orgId, kind, mode, to, success ? 1 : 0, providerId || null, error || null, key).run();
}
async function sendOne(env: Env, org: Organization, kind: "initial" | "followup", forceTest = false) {
  const settings = await env.DB.prepare("SELECT * FROM settings WHERE id=1").first<any>();
  const mode = forceTest || settings?.test_mode ? "test" : "live";
  const to = mode === "test" ? settings?.test_email || env.TEST_EMAIL : org.email;
  const day = dateKey(); const key = `${day}:${kind}:${org.id}:${mode}`;
  const already = await env.DB.prepare("SELECT id FROM send_logs WHERE idempotency_key=? AND success=1").bind(key).first();
  if (already) return { ok: false, skipped: "already_sent" };
  const unsubscribeUrl = mode === "test" ? `${env.APP_BASE_URL.replace(/\/$/, "")}/#test-mode` : `${env.APP_BASE_URL.replace(/\/$/, "")}/unsubscribe?token=${encodeURIComponent(org.unsubscribe_token)}`;
  try {
    const subject = kind === "initial" ? "Introducing JoblyHub – A Trusted Platform for Recruitment & Talent Discovery" : "Following up – JoblyHub Employer Partnership";
    const providerId = await sendProviderEmail(env, to, subject, emailBody(org, kind, unsubscribeUrl), key, settings?.from_email || env.FROM_EMAIL, settings?.from_email || env.FROM_EMAIL, settings?.from_name || env.FROM_NAME);
    await logSend(env, org.id, kind, mode, to, true, key, providerId);
    if (mode === "live") {
      const now = utcNow();
      if (kind === "initial") await env.DB.prepare("UPDATE organizations SET email_sent=1,date_sent=?,status='Sent',updated_at=? WHERE id=?").bind(now, now, org.id).run();
      else await env.DB.prepare("UPDATE organizations SET followup_sent=1,followup_sent_at=?,updated_at=? WHERE id=?").bind(now, now, org.id).run();
    }
    return { ok: true, mode, to, providerId };
  } catch (e) {
    const message = safe(e instanceof Error ? e.message : "Unknown send error", 500);
    await logSend(env, org.id, kind, mode, to, false, key, "", message);
    await env.DB.prepare("UPDATE organizations SET status='Failed',updated_at=? WHERE id=? AND ?='live'").bind(utcNow(), org.id, mode).run();
    await notify(env, "Email send failed", `${org.name} (${to}): ${message}`, "error");
    return { ok: false, error: message };
  }
}

async function runScheduled(env: Env, cron: string) {
  const settings = await env.DB.prepare("SELECT * FROM settings WHERE id=1").first<any>();
  if (!settings?.is_active) return { status: "paused" };
  const now = new Date(); const today = dateKey(now); const weekday = now.getUTCDay();
  if (weekday === 0 || weekday === 6) return { status: "weekend_skip" };
  if (cron === "0 8 * * 1-5") {
    const due = await env.DB.prepare(`SELECT * FROM organizations WHERE opted_out=0 AND approved=1 AND consent_status='written' AND response_received=0 AND email_sent=1 AND followup_sent=0 AND status='Sent' AND NOT EXISTS (SELECT 1 FROM send_logs t WHERE t.organization_id=organizations.id AND t.kind='followup' AND t.mode='test' AND t.success=1 AND substr(t.sent_at,1,10)=?) AND consent_record IS NOT NULL AND length(trim(consent_record))>0 AND planned_followup_date IS NOT NULL AND planned_followup_date<=? ORDER BY planned_followup_date,row_number LIMIT 3`).bind(today,today).all<Organization>();
    let sent = 0;
    if (settings.followups_enabled) for (const org of due.results) { const r = await sendOne(env, org, "followup"); if (r.ok) sent++; }
    await notify(env, "Morning workflow completed", `Follow-up check completed at 08:00. Follow-ups sent: ${sent}.`, "info");
    return { status: "morning_complete", followups_sent: sent };
  }
  const hour = now.getUTCHours(), minute = now.getUTCMinutes();
  if (hour < Number(settings.window_start_hour) || hour > Number(settings.window_end_hour) || (hour === Number(settings.window_end_hour) && minute > Number(settings.window_end_minute))) return { status: "outside_window" };
  const dailyMax = Math.max(1, Math.min(Number(settings.daily_max || env.DAILY_MAX || 15), 30));
  const currentMode = settings.test_mode ? "test" : "live";
  const count = await env.DB.prepare("SELECT COUNT(*) AS n FROM send_logs WHERE success=1 AND mode=? AND kind='initial' AND substr(sent_at,1,10)=?").bind(currentMode,today).first<{n:number}>();
  if ((count?.n || 0) >= dailyMax) return { status: "daily_limit_reached", sent_today: count?.n || 0 };
  const org = await env.DB.prepare(`SELECT * FROM organizations o WHERE o.email_sent=0 AND o.status IN ('Pending','Failed') AND o.approved=1 AND o.consent_status='written' AND o.consent_record IS NOT NULL AND length(trim(o.consent_record))>0 AND o.opted_out=0 AND o.email LIKE '%_@_%._%' AND NOT EXISTS (SELECT 1 FROM send_logs l WHERE l.organization_id=o.id AND l.kind='initial' AND l.mode='live' AND l.success=1) AND NOT EXISTS (SELECT 1 FROM send_logs t WHERE t.organization_id=o.id AND t.kind='initial' AND t.mode=? AND substr(t.sent_at,1,10)=?) ORDER BY COALESCE(o.planned_send_date,'0000-01-01'),o.row_number LIMIT 1`).bind(currentMode,today).first<Organization>();
  if (!org) return { status: "nothing_eligible" };
  const result = await sendOne(env, org, "initial");
  const current = await env.DB.prepare("SELECT COUNT(*) AS n FROM send_logs WHERE success=1 AND mode='live' AND kind='initial' AND substr(sent_at,1,10)=?").bind(today).first<{n:number}>();
  return { status: result.ok ? "sent" : result.skipped || "failed", sent_today: current?.n || 0, organization: org.name };
}

async function api(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url); const path = url.pathname; const method = req.method;
  if (path === "/api/health" && method === "GET") return json({ status: "ok", runtime: "Cloudflare Workers", database: "D1" });
  if (path === "/unsubscribe" && (method === "GET" || method === "POST")) {
    const token = url.searchParams.get("token") || "";
    if (!token) return html("<main><h1>Invalid unsubscribe link</h1><p>Please contact business@joblyhub.com.</p></main>", 400);
    const org = await env.DB.prepare("SELECT id,name FROM organizations WHERE unsubscribe_token=?").bind(token).first<{id:number,name:string}>();
    if (!org) return html("<main><h1>Link expired or invalid</h1><p>If you need help, email business@joblyhub.com.</p></main>", 404);
    if (method === "GET") return html(`<!doctype html><html><meta name="viewport" content="width=device-width"><title>Unsubscribe from JoblyHub</title><body style="font-family:system-ui;max-width:620px;margin:12vh auto;padding:24px;color:#123"><h1>Unsubscribe from JoblyHub emails</h1><p>Confirm that you want JoblyHub to stop sending outreach emails to ${escapeHtml(org.name)}.</p><form method="post" action="/unsubscribe?token=${encodeURIComponent(token)}"><button style="background:#006666;color:white;border:0;border-radius:7px;padding:12px 18px;font-weight:700;cursor:pointer">Confirm unsubscribe</button></form></body></html>`);
    await env.DB.prepare("UPDATE organizations SET opted_out=1,approved=0,status='Opted Out',updated_at=? WHERE id=?").bind(utcNow(), org.id).run();
    await notify(env, "Contact opted out", `${org.name} confirmed an unsubscribe request.`, "warning");
    return html("<!doctype html><html><meta name='viewport' content='width=device-width'><title>Unsubscribed</title><body style='font-family:system-ui;max-width:620px;margin:12vh auto;padding:24px;color:#123'><h1>You're unsubscribed</h1><p>JoblyHub has recorded your request and will not send further outreach emails to this address.</p></body></html>");
  }
  if (path.startsWith("/api/") && !await requireAdmin(req, env)) return adminError();

  if (path === "/api/dashboard" && method === "GET") {
    const counts = await env.DB.prepare(`SELECT COUNT(*) total, SUM(CASE WHEN email_sent=1 THEN 1 ELSE 0 END) sent, SUM(CASE WHEN status='Pending' THEN 1 ELSE 0 END) pending, SUM(CASE WHEN status='Bounced' THEN 1 ELSE 0 END) bounced, SUM(CASE WHEN response_received=1 THEN 1 ELSE 0 END) replied, SUM(CASE WHEN approved=1 AND consent_status='written' AND opted_out=0 AND email_sent=0 THEN 1 ELSE 0 END) eligible FROM organizations`).first<any>();
    const today = dateKey();
    const todaySent = await env.DB.prepare("SELECT COUNT(*) n FROM send_logs WHERE success=1 AND mode='live' AND kind='initial' AND substr(sent_at,1,10)=?").bind(today).first<any>();
    const recent = await env.DB.prepare("SELECT id,name,email,status,date_sent FROM organizations WHERE email_sent=1 ORDER BY date_sent DESC LIMIT 8").all();
    const settings = await env.DB.prepare("SELECT * FROM settings WHERE id=1").first<any>();
    const logs = await env.DB.prepare("SELECT id,title,message,level,is_read,created_at FROM notifications ORDER BY id DESC LIMIT 6").all();
    const sent = Number(counts?.sent || 0), bounced = Number(counts?.bounced || 0), replied = Number(counts?.replied || 0);
    return json({ counts: { ...counts, total:Number(counts?.total||0), sent, pending:Number(counts?.pending||0), bounced, replied, eligible:Number(counts?.eligible||0), delivery_rate:sent?Math.round((sent-bounced)/sent*1000)/10:0, bounce_rate:sent?Math.round(bounced/sent*1000)/10:0, reply_rate:sent?Math.round(replied/sent*1000)/10:0 }, today_sent:Number(todaySent?.n||0), recent:recent.results, settings, notifications:logs.results });
  }
  if (path === "/api/organizations" && method === "GET") {
    const page = Math.max(1, Number(url.searchParams.get("page") || 1)); const q = `%${safe(url.searchParams.get("q"), 120).replace(/[%_]/g, "")}%`; const status = safe(url.searchParams.get("status"), 80);
    const where = status ? "WHERE (name LIKE ? OR email LIKE ?) AND status=?" : "WHERE (name LIKE ? OR email LIKE ?)";
    const binds = status ? [q,q,status] : [q,q];
    const total = await env.DB.prepare(`SELECT COUNT(*) n FROM organizations ${where}`).bind(...binds).first<any>();
    const rows = await env.DB.prepare(`SELECT id,row_number,day_batch,name,sector,email,consent_status,approved,opted_out,email_sent,date_sent,status,response_received,followup_sent,planned_send_date,planned_followup_date FROM organizations ${where} ORDER BY row_number,id LIMIT 50 OFFSET ?`).bind(...binds,(page-1)*50).all();
    return json({ rows: rows.results, total:Number(total?.n||0), page, pages:Math.max(1,Math.ceil(Number(total?.n||0)/50)) });
  }
  if (path === "/api/import" && method === "POST") {
    const form = await req.formData(); const file = form.get("file");
    if (!(file instanceof File)) return json({ error: "Upload a CSV or XLSX file." }, 400);
    if (file.size > 5 * 1024 * 1024) return json({ error: "Maximum file size is 5 MB." }, 413);
    const filename = file.name.toLowerCase();
    let parsed: ReturnType<typeof parseCsv>;
    try {
      if (filename.endsWith(".xlsx")) parsed = parseXlsx(await file.arrayBuffer());
      else if (filename.endsWith(".csv") || filename.endsWith(".txt")) parsed = parseCsv(await file.text());
      else return json({ error: "Unsupported file type. Upload .csv or .xlsx." }, 415);
    } catch (e) {
      return json({ error: safe(e instanceof Error ? e.message : "Could not parse the uploaded file.", 500) }, 400);
    }
    let imported = 0, skipped = 0, errors: string[] = [];
    const statements: D1PreparedStatement[] = [];
    for (const r of parsed) {
      if (!r.name || !emailValid(r.email)) { skipped++; if (errors.length < 8) errors.push(`Row ${r.row_number}: missing organization name or valid email.`); continue; }
      const token = crypto.randomUUID().replace(/-/g, "") + crypto.randomUUID().replace(/-/g, "");
      statements.push(env.DB.prepare(`INSERT INTO organizations(row_number,day_batch,day_number,batch_position,planned_send_date,planned_followup_date,name,sector,email,email_normalized,is_ngo,consent_status,consent_record,approved,opted_out,unsubscribe_token,email_sent,date_sent,status,response_received,response_sentiment,responded_at,followup_sent,followup_sent_at,account_created,job_posted,became_paid,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(email_normalized) DO UPDATE SET row_number=excluded.row_number,day_batch=excluded.day_batch,day_number=excluded.day_number,batch_position=excluded.batch_position,planned_send_date=COALESCE(excluded.planned_send_date,organizations.planned_send_date),planned_followup_date=COALESCE(excluded.planned_followup_date,organizations.planned_followup_date),name=excluded.name,sector=excluded.sector,is_ngo=excluded.is_ngo,consent_status=CASE WHEN excluded.consent_status='written' THEN 'written' WHEN excluded.consent_status='denied' THEN 'denied' ELSE organizations.consent_status END,consent_record=COALESCE(excluded.consent_record,organizations.consent_record),approved=CASE WHEN excluded.opted_out=1 OR excluded.consent_status='denied' THEN 0 WHEN excluded.approved=1 THEN 1 ELSE organizations.approved END,opted_out=MAX(organizations.opted_out,excluded.opted_out),email_sent=MAX(organizations.email_sent,excluded.email_sent),date_sent=COALESCE(organizations.date_sent,excluded.date_sent),status=CASE WHEN organizations.email_sent=1 THEN organizations.status WHEN excluded.email_sent=1 THEN COALESCE(excluded.status,'Sent') ELSE organizations.status END,response_received=MAX(organizations.response_received,excluded.response_received),response_sentiment=COALESCE(excluded.response_sentiment,organizations.response_sentiment),responded_at=COALESCE(organizations.responded_at,excluded.responded_at),followup_sent=MAX(organizations.followup_sent,excluded.followup_sent),followup_sent_at=COALESCE(organizations.followup_sent_at,excluded.followup_sent_at),account_created=MAX(organizations.account_created,excluded.account_created),job_posted=MAX(organizations.job_posted,excluded.job_posted),became_paid=MAX(organizations.became_paid,excluded.became_paid),notes=COALESCE(excluded.notes,organizations.notes),updated_at=CURRENT_TIMESTAMP`).bind(r.row_number,r.day_batch,r.day_number,r.batch_position,r.planned_send_date,r.planned_followup_date,r.name,r.sector,r.email,r.email,r.is_ngo,r.consent_status,r.consent_record,r.approved,r.opted_out,token,r.email_sent,r.date_sent,r.status,r.response_received,r.response_sentiment,r.responded_at,r.followup_sent,r.followup_sent_at,r.account_created,r.job_posted,r.became_paid,r.notes));
    }
    for (let i = 0; i < statements.length; i += 50) {
      const chunk = statements.slice(i, i + 50);
      try { await env.DB.batch(chunk); imported += chunk.length; }
      catch (e) { skipped += chunk.length; if (errors.length < 8) errors.push(`A batch of ${chunk.length} valid rows failed to import. Check CSV values and try again.`); console.error("CSV import batch failed", e); }
    }
    await notify(env,"Lead import completed",`CSV import processed ${parsed.length} rows; upserted ${imported}; skipped ${skipped}. Consent/approval are not inferred if absent.`, errors.length ? "warning":"success");
    return json({ processed:parsed.length, imported, skipped, errors, note:"Rows without explicit written-consent and approval fields remain ineligible for automatic sending." });
  }
  const orgAction = path.match(/^\/api\/organizations\/(\d+)$/);
  if (orgAction && method === "PATCH") {
    const id = Number(orgAction[1]); const body:any = await req.json();
    const existing = await env.DB.prepare("SELECT id FROM organizations WHERE id=?").bind(id).first(); if (!existing) return json({error:"Lead not found."},404);
    const fields: string[] = [], values:any[] = [];
    if (body.approved !== undefined) { fields.push("approved=?"); values.push(bool(body.approved)?1:0); }
    if (body.consent_status !== undefined) { const cs = ["unknown","written","denied"].includes(body.consent_status) ? body.consent_status : "unknown"; fields.push("consent_status=?"); values.push(cs); if (cs === "denied") fields.push("approved=0"); }
    if (body.consent_record !== undefined) { fields.push("consent_record=?"); values.push(safe(body.consent_record,1000)||null); }
    if (body.opted_out !== undefined) { fields.push("opted_out=?"); values.push(bool(body.opted_out)?1:0); if(bool(body.opted_out)){fields.push("approved=0");} }
    if (body.status !== undefined) { const statuses=["Pending","Sent","Bounced","Out of Office","Replied - Positive","Replied - Negative","Mailbox Full","Failed","Opted Out"]; if(!statuses.includes(body.status))return json({error:"Invalid status."},400); fields.push("status=?"); values.push(body.status); if(String(body.status).startsWith("Replied")){fields.push("response_received=1");fields.push("responded_at=COALESCE(responded_at,?)");values.push(utcNow());} if(body.status==="Opted Out"){fields.push("opted_out=1");fields.push("approved=0");} }
    if (!fields.length) return json({error:"No supported fields provided."},400);
    fields.push("updated_at=?"); values.push(utcNow(),id);
    await env.DB.prepare(`UPDATE organizations SET ${fields.join(",")} WHERE id=?`).bind(...values).run();
    return json({ok:true});
  }
  if (path === "/api/settings" && method === "GET") return json(await env.DB.prepare("SELECT * FROM settings WHERE id=1").first());
  if (path === "/api/settings" && method === "PUT") {
    const b:any = await req.json();
    const max = Math.max(1, Math.min(Number(b.daily_max || 15), 30));
    const min = Math.max(0, Math.min(Number(b.daily_min || 10), max));
    const start = Math.max(0, Math.min(Number(b.window_start_hour ?? 9), 23));
    const end = Math.max(start, Math.min(Number(b.window_end_hour ?? 16), 23));
    await env.DB.prepare(`UPDATE settings SET is_active=?,test_mode=?,test_email=?,from_name=?,from_email=?,daily_min=?,daily_max=?,window_start_hour=?,window_end_hour=?,window_end_minute=?,followups_enabled=?,updated_at=? WHERE id=1`).bind(bool(b.is_active)?1:0,bool(b.test_mode)?1:0,emailValid(safe(b.test_email,254))?safe(b.test_email,254):env.TEST_EMAIL,safe(b.from_name,100)||env.FROM_NAME,emailValid(safe(b.from_email,254))?safe(b.from_email,254):env.FROM_EMAIL,min,max,start,end,Math.max(0,Math.min(Number(b.window_end_minute ?? 30),59)),bool(b.followups_enabled)?1:0,utcNow()).run();
    await notify(env,"Settings updated",`Assistant ${bool(b.is_active)?"enabled":"paused"}; test mode ${bool(b.test_mode)?"ON":"OFF"}.`,"info");
    return json({ok:true});
  }
  if (path === "/api/run-test" && method === "POST") {
    const org = await env.DB.prepare("SELECT * FROM organizations WHERE email_sent=0 AND opted_out=0 ORDER BY id LIMIT 1").first<Organization>();
    if (!org) return json({error:"Import at least one lead first."},400);
    const result = await sendOne(env, org, "initial", true);
    return json(result, result.ok ? 200 : 502);
  }
  if (path === "/api/followups" && method === "GET") {
    const rows = await env.DB.prepare(`SELECT id,name,email,sector,status,planned_followup_date,email_sent,followup_sent,consent_status,consent_record,approved,opted_out FROM organizations WHERE email_sent=1 AND followup_sent=0 AND status='Sent' AND planned_followup_date IS NOT NULL ORDER BY planned_followup_date,row_number LIMIT 250`).all();
    return json({ rows: rows.results, today: dateKey() });
  }
  if (path === "/api/reports" && method === "GET") {
    const now = new Date(); const monday = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())); monday.setUTCDate(monday.getUTCDate() - ((monday.getUTCDay()+6)%7)); const weekStart = monday.toISOString();
    const summary = await env.DB.prepare(`SELECT COUNT(*) sent,SUM(CASE WHEN status='Bounced' THEN 1 ELSE 0 END) bounced,SUM(response_received) replies,SUM(account_created) accounts_created,SUM(job_posted) jobs_posted,SUM(became_paid) paid_customers FROM organizations WHERE date_sent>=?`).bind(weekStart).first<any>();
    const sectors = await env.DB.prepare(`SELECT COALESCE(sector,'Uncategorised') sector,COUNT(*) sent,SUM(response_received) replies FROM organizations WHERE date_sent>=? GROUP BY sector ORDER BY replies DESC LIMIT 10`).bind(weekStart).all();
    return json({ week_start: weekStart.slice(0,10), summary: summary || {}, sectors: sectors.results });
  }
  if (path === "/api/activity" && method === "GET") {
    const logs = await env.DB.prepare("SELECT l.*,o.name AS organization_name FROM send_logs l LEFT JOIN organizations o ON o.id=l.organization_id ORDER BY l.id DESC LIMIT 100").all();
    const notifications = await env.DB.prepare("SELECT * FROM notifications ORDER BY id DESC LIMIT 100").all();
    return json({ logs:logs.results, notifications:notifications.results });
  }
  if (path === "/api/notifications/read" && method === "POST") { await env.DB.prepare("UPDATE notifications SET is_read=1").run(); return json({ok:true}); }
  if (path === "/api/analytics" && method === "GET") {
    const sectors = await env.DB.prepare(`SELECT COALESCE(sector,'Uncategorised') sector,COUNT(*) sent,SUM(response_received) replies FROM organizations WHERE email_sent=1 GROUP BY sector ORDER BY replies DESC LIMIT 12`).all();
    const funnel = await env.DB.prepare(`SELECT COUNT(*) organizations,SUM(email_sent) sent,SUM(response_received) replies,SUM(account_created) accounts_created,SUM(job_posted) jobs_posted,SUM(became_paid) paid_customers FROM organizations`).first();
    const daily = await env.DB.prepare(`SELECT substr(date_sent,1,10) day,COUNT(*) sent,SUM(response_received) replies,SUM(CASE WHEN status='Bounced' THEN 1 ELSE 0 END) bounces FROM organizations WHERE date_sent>=date('now','-13 day') GROUP BY substr(date_sent,1,10) ORDER BY day`).all();
    return json({sectors:sectors.results,funnel,daily:daily.results});
  }
  if (path === "/api/cron-status" && method === "GET") {
    const today = dateKey(); const sent = await env.DB.prepare("SELECT COUNT(*) n FROM send_logs WHERE success=1 AND mode='live' AND kind='initial' AND substr(sent_at,1,10)=?").bind(today).first<any>();
    const settings = await env.DB.prepare("SELECT is_active,test_mode,daily_max,window_start_hour,window_end_hour,window_end_minute FROM settings WHERE id=1").first();
    return json({today,sent:Number(sent?.n||0),settings,timezone:"Africa/Accra (UTC+0)",triggers:["08:00 weekdays: follow-up check","09:00–16:30 weekdays: paced outreach"]});
  }
  return json({error:"Not found"},404);
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname.startsWith("/api/") || url.pathname === "/unsubscribe") {
      try { await ensureSchema(env); return await api(request, env); }
      catch (e) { console.error("Request failed", e); return json({error: e instanceof Error ? e.message : "Internal error"},500); }
    }
    const asset = await env.ASSETS.fetch(request);
    if (asset.status !== 404) return asset;
    if (request.method === "GET") return env.ASSETS.fetch(new Request(new URL("/index.html", request.url), { headers: request.headers, method: "GET" }));
    return json({error:"Not found"},404);
  },
  async scheduled(event: ScheduledController, env: Env, ctx: ExecutionContext) {
    ctx.waitUntil((async () => {
      try { await ensureSchema(env); const result = await runScheduled(env, event.cron); console.log("Sasa scheduled run", event.cron, result); }
      catch (e) { console.error("Scheduled run failed", e); await notify(env,"Scheduled run failed",safe(e instanceof Error?e.message:"Unknown error",1000),"error"); }
    })());
  }
};

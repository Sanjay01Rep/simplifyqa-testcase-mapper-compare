/**
 * Microsoft Graph helpers: device-code login and Reply-All draft in Outlook on the web.
 * To/CC are never set here — createReplyAll keeps the live thread recipients.
 */

const fs = require("fs");
const path = require("path");
const { getEmailChain, buildStatusEmail, formatMailDate, todayMailDate } = require("./emailChain");

const ROOT = path.join(__dirname, "../..");
const TOKEN_PATH = path.join(ROOT, "config", "outlook-graph-token.json");
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = "offline_access User.Read Mail.ReadWrite";
const SIMPLE_ATTACH_MAX = 2.5 * 1024 * 1024;

let pendingDevice = null;

function clientId() {
  return String(process.env.MICROSOFT_CLIENT_ID || "").trim();
}

function tenantId() {
  return String(process.env.MICROSOFT_TENANT_ID || "common").trim() || "common";
}

function loginBase() {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenantId())}/oauth2/v2.0`;
}

function outlookConfigured() {
  return Boolean(clientId());
}

function readTokenFile() {
  try {
    if (!fs.existsSync(TOKEN_PATH)) return null;
    const raw = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"));
    if (!raw || !raw.refresh_token) return null;
    return raw;
  } catch {
    return null;
  }
}

function writeTokenFile(token) {
  fs.mkdirSync(path.dirname(TOKEN_PATH), { recursive: true });
  const payload = {
    access_token: String(token.access_token || ""),
    refresh_token: String(token.refresh_token || ""),
    expires_at: Number(token.expires_at) || 0,
    account: String(token.account || ""),
  };
  fs.writeFileSync(TOKEN_PATH, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

function clearTokenFile() {
  pendingDevice = null;
  try {
    if (fs.existsSync(TOKEN_PATH)) fs.unlinkSync(TOKEN_PATH);
  } catch {
    /* ignore */
  }
}

function formBody(obj) {
  return Object.entries(obj)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");
}

async function exchangeDeviceCode(deviceCode) {
  const res = await fetch(`${loginBase()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: formBody({
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId(),
      device_code: deviceCode,
    }),
  });
  const data = await res.json();
  return { ok: res.ok, data };
}

async function refreshAccessToken(stored) {
  const res = await fetch(`${loginBase()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: formBody({
      grant_type: "refresh_token",
      client_id: clientId(),
      refresh_token: stored.refresh_token,
      scope: SCOPES,
    }),
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error_description || data.error || "Outlook token refresh failed.");
  }
  const next = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || stored.refresh_token,
    expires_at: Date.now() + Number(data.expires_in || 3600) * 1000 - 60 * 1000,
    account: stored.account || "",
  };
  if (!next.account) {
    try {
      next.account = await readAccount(next.access_token);
    } catch {
      next.account = stored.account || "";
    }
  }
  writeTokenFile(next);
  return next;
}

async function readAccount(accessToken) {
  const res = await fetch(`${GRAPH}/me?$select=userPrincipalName,mail,displayName`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await res.json();
  if (!res.ok) return "";
  return String(data.userPrincipalName || data.mail || data.displayName || "").trim();
}

async function getAccessToken() {
  const stored = readTokenFile();
  if (!stored) return null;
  if (stored.access_token && stored.expires_at && Date.now() < stored.expires_at) {
    return stored.access_token;
  }
  const next = await refreshAccessToken(stored);
  return next.access_token;
}

function outlookStatus() {
  const stored = readTokenFile();
  return {
    configured: outlookConfigured(),
    connected: Boolean(stored && stored.refresh_token),
    account: stored && stored.account ? stored.account : "",
    tenant: tenantId(),
  };
}

async function startDeviceLogin() {
  if (!outlookConfigured()) {
    throw new Error(
      "Set MICROSOFT_CLIENT_ID in .env (Azure app with public client + Mail.ReadWrite) to connect Outlook on the web."
    );
  }
  const res = await fetch(`${loginBase()}/devicecode`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: formBody({
      client_id: clientId(),
      scope: SCOPES,
    }),
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error_description || data.error || "Could not start Outlook sign-in.");
  }
  pendingDevice = {
    device_code: data.device_code,
    interval: Number(data.interval) || 5,
    expires_at: Date.now() + Number(data.expires_in || 900) * 1000,
  };
  return {
    userCode: data.user_code,
    verificationUrl: data.verification_uri || "https://microsoft.com/devicelogin",
    message: data.message || `Go to ${data.verification_uri} and enter ${data.user_code}`,
    expiresIn: Number(data.expires_in) || 900,
  };
}

async function pollDeviceLogin() {
  if (!pendingDevice) {
    return { pending: false, connected: outlookStatus().connected };
  }
  if (Date.now() > pendingDevice.expires_at) {
    pendingDevice = null;
    throw new Error("Outlook sign-in timed out. Connect again.");
  }
  const { ok, data } = await exchangeDeviceCode(pendingDevice.device_code);
  if (ok && data.access_token) {
    const account = await readAccount(data.access_token);
    writeTokenFile({
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + Number(data.expires_in || 3600) * 1000 - 60 * 1000,
      account,
    });
    pendingDevice = null;
    return { pending: false, connected: true, account };
  }
  if (data.error === "authorization_pending" || data.error === "slow_down") {
    return { pending: true, connected: false };
  }
  pendingDevice = null;
  throw new Error(data.error_description || data.error || "Outlook sign-in failed.");
}

async function graphJson(method, urlPath, body, token) {
  const res = await fetch(`${GRAPH}${urlPath}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body == null ? undefined : JSON.stringify(body),
  });
  const text = await res.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }
  if (!res.ok) {
    const msg =
      (data.error && data.error.message) ||
      data.error_description ||
      `Microsoft Graph ${res.status}`;
    const err = new Error(msg);
    err.status = res.status;
    throw err;
  }
  return data;
}

async function findLatestThreadMessage(token, search) {
  const q = String(search || "").trim();
  if (!q) return null;
  const url = `/me/messages?$search=${encodeURIComponent(`"${q}"`)}&$top=15&$select=id,subject,receivedDateTime,conversationId,webLink,isDraft`;
  const data = await graphJson("GET", url, null, token);
  const rows = (data.value || [])
    .filter((m) => m && m.id && !m.isDraft)
    .sort(
      (a, b) =>
        new Date(b.receivedDateTime || 0).getTime() - new Date(a.receivedDateTime || 0).getTime()
    );
  return rows[0] || null;
}

async function addAttachment(token, messageId, excelPath) {
  const name = path.basename(excelPath);
  const buf = fs.readFileSync(excelPath);
  if (buf.length <= SIMPLE_ATTACH_MAX) {
    await graphJson(
      "POST",
      `/me/messages/${encodeURIComponent(messageId)}/attachments`,
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        contentBytes: buf.toString("base64"),
      },
      token
    );
    return;
  }
  const session = await graphJson(
    "POST",
    `/me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`,
    {
      AttachmentItem: {
        attachmentType: "file",
        name,
        size: buf.length,
      },
    },
    token
  );
  const uploadUrl = session.uploadUrl;
  if (!uploadUrl) throw new Error("Outlook did not return an attachment upload URL.");
  const chunk = 4 * 1024 * 1024;
  for (let start = 0; start < buf.length; start += chunk) {
    const end = Math.min(start + chunk, buf.length) - 1;
    const slice = buf.subarray(start, end + 1);
    const put = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Length": String(slice.length),
        "Content-Range": `bytes ${start}-${end}/${buf.length}`,
      },
      body: slice,
    });
    if (!put.ok && put.status !== 201 && put.status !== 200) {
      const t = await put.text();
      throw new Error(`Outlook attachment upload failed: HTTP ${put.status} ${t.slice(0, 200)}`);
    }
  }
}

async function createOutlookChainDraft({
  templateChoice,
  templateRelative,
  excelPath,
  sheetName,
  compareSheetName,
}) {
  const chain = getEmailChain(templateChoice, templateRelative);
  if (!chain) {
    throw new Error("This template has no email chain configured yet (SIT and UAT 1 are available).");
  }
  if (!chain.search) {
    throw new Error(`Email chain for template ${chain.choice} is missing a mailbox search subject.`);
  }
  if (!outlookConfigured()) {
    throw new Error("Set MICROSOFT_CLIENT_ID in .env to draft in Outlook on the web.");
  }
  const token = await getAccessToken();
  if (!token) {
    throw new Error("Connect Outlook on the web first (one-time Microsoft sign-in).");
  }

  const asOfDate = formatMailDate(sheetName) || todayMailDate();
  const compareDate = formatMailDate(compareSheetName) || asOfDate;
  const built = buildStatusEmail({ chain, asOfDate, compareDate });

  const latest = await findLatestThreadMessage(token, chain.search);
  if (!latest) {
    throw new Error(
      `No existing mail found for “${chain.search}”. Open that thread once in Outlook on the web, then try again. To and CC stay on that conversation.`
    );
  }

  const created = await graphJson(
    "POST",
    `/me/messages/${encodeURIComponent(latest.id)}/createReplyAll`,
    {},
    token
  );
  const draftId = created.id;
  if (!draftId) {
    throw new Error("Outlook did not return a Reply All draft.");
  }
  let draft = created;
  if (!draft.body || !draft.body.content) {
    draft = await graphJson(
      "GET",
      `/me/messages/${encodeURIComponent(draftId)}?$select=id,subject,body,webLink`,
      null,
      token
    );
  }
  const quoted = (draft.body && draft.body.content) || "";
  const html = `${built.html}${quoted}`;
  const updated = await graphJson(
    "PATCH",
    `/me/messages/${encodeURIComponent(draftId)}`,
    {
      subject: built.subject,
      body: { contentType: "HTML", content: html },
    },
    token
  );

  if (chain.attachExcel) {
    const abs = path.isAbsolute(excelPath) ? excelPath : path.join(ROOT, excelPath);
    if (!fs.existsSync(abs)) {
      throw new Error(`Generated Excel was not found: ${abs}`);
    }
    await addAttachment(token, draftId, abs);
  }

  const fresh = await graphJson(
    "GET",
    `/me/messages/${encodeURIComponent(draftId)}?$select=id,subject,webLink,isDraft`,
    null,
    token
  );
  return {
    id: fresh.id || draftId,
    subject: fresh.subject || updated.subject || built.subject,
    webLink: fresh.webLink || updated.webLink || "",
    asOfDate: built.asOfDate,
    compareDate: built.compareDate,
    threadSubject: latest.subject || "",
  };
}

module.exports = {
  TOKEN_PATH,
  outlookConfigured,
  outlookStatus,
  startDeviceLogin,
  pollDeviceLogin,
  clearTokenFile,
  createOutlookChainDraft,
};

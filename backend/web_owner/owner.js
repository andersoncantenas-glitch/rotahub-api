const API_BASE = "/api/v1";
const TOKEN_KEY = "rotahub_owner_token";

const state = {
  token: sessionStorage.getItem(TOKEN_KEY) || "",
  user: null,
  data: null,
  selectedCompanyId: null,
  currentTab: "dashboard",
};

const el = (id) => document.getElementById(id);
const clean = (value) => String(value ?? "").trim();
const money = (value) => Number(value || 0).toLocaleString("pt-BR", {style: "currency", currency: "BRL"});
const number = (value) => Number(value || 0).toLocaleString("pt-BR", {maximumFractionDigits: 2});
const PAYMENT_GRACE_DAYS = 3;
const LATE_FINE_PERCENT = 2;
const MONTHLY_INTEREST_PERCENT = 1;
const statusLabels = {
  active: "Ativa",
  suspended: "Suspensa",
  inactive: "Inativa",
  cancelled: "Cancelada",
  pending: "Pendente",
  approved: "Aprovada",
  rejected: "Recusada",
  upgrade: "Upgrade",
  downgrade: "Downgrade",
  change: "Troca",
  paid: "Pago",
  overdue: "Em atraso",
  trialing: "Demonstração",
  novo: "Pendente",
  aprovado: "Aprovada",
  recusado: "Recusada",
};
const escapeHtml = (value) => clean(value)
  .replaceAll("&", "&amp;")
  .replaceAll("<", "&lt;")
  .replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;")
  .replaceAll("'", "&#039;");

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  if (state.token) loadOwnerSession();
});

function bindEvents() {
  el("loginForm").addEventListener("submit", login);
  el("logoutButton").addEventListener("click", logout);
  el("refreshButton").addEventListener("click", loadDashboard);
  el("planForm").addEventListener("submit", changePlan);
  el("adminAccessButton").addEventListener("click", resetAdminAccess);
  el("paymentForm").addEventListener("submit", createPayment);
  el("paymentCompanySelect").addEventListener("change", changePaymentCompany);
  el("paymentAmount").addEventListener("input", updatePaymentLatePreview);
  el("paymentDue").addEventListener("change", updatePaymentLatePreview);
  el("overdueButton").addEventListener("click", runOverdue);
  el("companySearch").addEventListener("input", renderCompanyTable);
  el("companyStatusFilter").addEventListener("change", renderCompanyTable);
  el("leadSearch").addEventListener("input", renderLeadTable);
  el("leadStatusFilter").addEventListener("change", renderLeadTable);
  el("planRequestSearch").addEventListener("input", renderPlanRequestTable);
  el("planRequestStatusFilter").addEventListener("change", renderPlanRequestTable);
  el("copyAccessMessage").addEventListener("click", copyAccessMessage);
  el("accessModalClose").addEventListener("click", closeAccessModal);
  el("accessModalCloseTop").addEventListener("click", closeAccessModal);
  el("accessModal").addEventListener("click", (event) => {
    if (event.target === el("accessModal")) closeAccessModal();
  });
  document.querySelectorAll("[data-tab]").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
}

async function api(path, options = {}) {
  const headers = new Headers(options.headers || {});
  if (!(options.body instanceof FormData)) headers.set("Content-Type", "application/json");
  if (state.token) headers.set("Authorization", `Bearer ${state.token}`);
  const response = await fetch(`${API_BASE}${path}`, {...options, headers});
  if (response.status === 401) {
    logout();
    throw new Error("Sessão expirada.");
  }
  if (!response.ok) {
    let detail = "Requisição recusada.";
    try {
      const data = await response.json();
      detail = data.detail || detail;
    } catch (_error) {
      detail = await response.text() || detail;
    }
    throw new Error(Array.isArray(detail) ? detail.map((item) => item.msg).join("; ") : detail);
  }
  return response.json();
}

async function login(event) {
  event.preventDefault();
  el("loginError").textContent = "";
  const body = new URLSearchParams();
  body.set("username", clean(el("username").value));
  body.set("password", el("password").value);
  try {
    const response = await fetch(`${API_BASE}/auth/login`, {
      method: "POST",
      headers: {"Content-Type": "application/x-www-form-urlencoded"},
      body,
    });
    if (!response.ok) throw new Error("Usuário ou senha inválidos.");
    const data = await response.json();
    state.token = data.access_token || "";
    sessionStorage.setItem(TOKEN_KEY, state.token);
    await loadOwnerSession();
  } catch (error) {
    logout();
    el("loginError").textContent = error.message;
  }
}

async function loadOwnerSession() {
  try {
    state.user = await api("/users/me");
    await loadDashboard();
    el("loginView").classList.add("hidden");
    el("ownerShell").classList.remove("hidden");
    el("currentUser").textContent = `${state.user.nome || state.user.username} (${state.user.permissoes})`;
  } catch (error) {
    el("loginError").textContent = error.message;
  }
}

function logout() {
  state.token = "";
  state.user = null;
  sessionStorage.removeItem(TOKEN_KEY);
  el("ownerShell").classList.add("hidden");
  el("loginView").classList.remove("hidden");
}

function toast(message, error = false) {
  const box = el("toast");
  box.textContent = message;
  box.style.background = error ? "#b42318" : "#102033";
  box.classList.remove("hidden");
  window.clearTimeout(toast.timer);
  toast.timer = window.setTimeout(() => box.classList.add("hidden"), 3200);
}

async function loadDashboard() {
  try {
    const query = state.selectedCompanyId ? `?company_id=${encodeURIComponent(state.selectedCompanyId)}` : "";
    state.data = await api(`/saas-admin/dashboard${query}`);
    const selected = state.data.company || (state.data.companies || [])[0] || null;
    state.selectedCompanyId = selected ? Number(selected.id) : null;
    render();
  } catch (error) {
    toast(error.message, true);
  }
}

function switchTab(tab) {
  state.currentTab = tab;
  document.querySelectorAll("[data-tab]").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  document.querySelectorAll(".tab").forEach((section) => section.classList.add("hidden"));
  el(`${tab}Tab`).classList.remove("hidden");
}

function currentCompany() {
  const companies = state.data?.companies || [];
  return companies.find((item) => Number(item.id) === Number(state.selectedCompanyId)) || state.data?.company || null;
}

function companyDisplayName(company) {
  return company?.name || company?.nome || company?.razao_social || company?.legal_name || `Empresa #${company?.id || ""}`;
}

function companyOptionLabel(company) {
  const parts = [
    `#${company.id}`,
    companyDisplayName(company),
    company.document || company.cnpj || company.cpf || "",
    company.plan_code || company.plano || "",
    statusLabel(company.status),
  ].map(clean).filter(Boolean);
  return parts.join(" | ");
}

function planForCompany(company) {
  const plans = state.data?.plans || [];
  const code = clean(company?.plan_code || company?.plano).toLowerCase();
  return plans.find((plan) => clean(plan.code).toLowerCase() === code) || null;
}

function planMonthlyPrice(plan) {
  return Number(plan?.monthly_price ?? plan?.price ?? plan?.valor ?? 0);
}

function formatMoneyInput(value) {
  return Number(value || 0).toLocaleString("pt-BR", {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function parseDateOnly(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;
  const [year, month, day] = raw.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function daysAfterGrace(dueDateValue) {
  const due = parseDateOnly(dueDateValue);
  if (!due) return 0;
  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const msPerDay = 24 * 60 * 60 * 1000;
  const daysLate = Math.floor((todayOnly.getTime() - due.getTime()) / msPerDay);
  return Math.max(0, daysLate - PAYMENT_GRACE_DAYS);
}

function adjustedPaymentAmount(baseAmount, dueDateValue) {
  const base = Number(baseAmount || 0);
  const chargeableDays = daysAfterGrace(dueDateValue);
  if (!base || chargeableDays <= 0) return base;
  const fine = base * (LATE_FINE_PERCENT / 100);
  const interest = base * (MONTHLY_INTEREST_PERCENT / 100) * (chargeableDays / 30);
  return Number((base + fine + interest).toFixed(2));
}

function statusLabel(value) {
  const key = clean(value).toLowerCase();
  return statusLabels[key] || clean(value) || "-";
}

function statusBadge(value) {
  const key = clean(value).toLowerCase();
  return `<span class="status-badge ${escapeHtml(key)}">${escapeHtml(statusLabel(key))}</span>`;
}

function dateTime(value) {
  const raw = clean(value);
  if (!raw) return "-";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [year, month, day] = raw.split("-");
    return `${day}/${month}/${year}`;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-BR", {dateStyle: "short", timeStyle: "short"});
}

function parseMoney(value) {
  const raw = clean(value).replace(/\s+/g, "").replace(/^R\$/i, "");
  if (!raw) return NaN;
  const normalized = raw.includes(",") ? raw.replaceAll(".", "").replace(",", ".") : raw;
  return Number(normalized);
}

function accessMessage({companyName = "", username = "", password = "", trialDays = "", trialEnd = ""} = {}) {
  const lines = [
    `Olá${companyName ? `, ${companyName}` : ""}.`,
    "",
    "Seu acesso ao RotaHub foi liberado.",
    "",
    "Link de acesso:",
    `${window.location.origin}/app/index.html`,
    "",
    "Dados de acesso:",
    `Usuário: ${username}`,
    `Senha temporária: ${password}`,
  ];
  if (trialDays || trialEnd) {
    lines.push("", "Demonstração:");
    if (trialDays) lines.push(`${trialDays} dia(s) de uso liberado.`);
    if (trialEnd) lines.push(`Válida até: ${trialEnd}.`);
  }
  lines.push(
    "",
    "Por segurança, recomendamos trocar a senha no primeiro acesso.",
    "",
    "Qualquer dúvida, entre em contato com a equipe RotaHub."
  );
  return lines.join("\n");
}

function absoluteUrl(path) {
  const raw = clean(path);
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  return `${window.location.origin}${raw.startsWith("/") ? raw : `/${raw}`}`;
}

function paymentCompanyName(payment) {
  const company = currentCompany() || {};
  return company.name || company.nome || company.razao_social || `Empresa #${payment.company_id || state.selectedCompanyId || ""}`;
}

function boletoMessage(payment) {
  const link = absoluteUrl(payment.boleto_pdf_url);
  const lines = [
    `Olá, ${paymentCompanyName(payment)}.`,
    "",
    "Segue boleto da sua assinatura RotaHub.",
    "",
    `Valor: ${money(payment.amount || 0)}`,
    `Vencimento: ${dateTime(payment.due_date)}`,
  ];
  if (payment.boleto_digitable_line) lines.push(`Linha digitável: ${payment.boleto_digitable_line}`);
  if (link) lines.push("", `Boleto: ${link}`);
  lines.push("", "Após o pagamento, envie o comprovante para registrarmos a baixa.");
  return lines.join("\n");
}

async function copyText(text, successMessage) {
  try {
    await navigator.clipboard.writeText(text);
    toast(successMessage);
  } catch (_error) {
    window.prompt("Copie a mensagem", text);
  }
}

function openAccessModal(payload) {
  el("accessMessage").value = accessMessage(payload);
  el("copyAccessStatus").textContent = "";
  el("accessModal").classList.remove("hidden");
  el("accessMessage").focus();
  el("accessMessage").select();
}

function closeAccessModal() {
  el("accessModal").classList.add("hidden");
}

async function copyAccessMessage() {
  const message = el("accessMessage").value;
  try {
    await navigator.clipboard.writeText(message);
    el("copyAccessStatus").textContent = "Mensagem copiada.";
  } catch (_error) {
    el("accessMessage").focus();
    el("accessMessage").select();
    document.execCommand("copy");
    el("copyAccessStatus").textContent = "Mensagem selecionada para copiar.";
  }
}

function render() {
  const data = state.data || {};
  const companies = data.companies || [];
  const payments = data.payments || [];
  const audit = data.audit_logs || [];
  const plans = data.plans || [];
  const leads = data.signup_leads || [];
  const planRequests = data.plan_change_requests || [];
  const company = currentCompany();
  const pendingLeads = leads.filter((item) => clean(item.status).toLowerCase() === "novo").length;
  const pendingPlanRequests = planRequests.filter((item) => clean(item.status).toLowerCase() === "pending").length;

  el("metrics").innerHTML = [
    ["Empresas", companies.length],
    ["Empresas ativas", companies.filter((item) => clean(item.status).toLowerCase() === "active").length],
    ["Cadastros pendentes", pendingLeads],
    ["Trocas de plano", pendingPlanRequests],
    ["Cobranças da empresa", payments.length],
  ].map(([label, value]) => `<article class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`).join("");
  el("leadNavCount").textContent = pendingLeads;
  el("planRequestNavCount").textContent = pendingPlanRequests;

  renderCompanyTable();
  renderLeadTable();
  renderPlanRequestTable();
  renderSelectedCompany(company, plans, data.usage || {}, data.features || {});
  renderPayments(payments);
  renderAudit(audit);
}

function renderLeadTable() {
  const leads = state.data?.signup_leads || [];
  const plans = state.data?.plans || [];
  const term = clean(el("leadSearch").value).toLowerCase();
  const status = clean(el("leadStatusFilter").value).toLowerCase();
  const filtered = leads.filter((lead) => {
    const searchable = [lead.company, lead.name, lead.document, lead.email, lead.phone, lead.plan_code]
      .map(clean).join(" ").toLowerCase();
    return (!status || clean(lead.status).toLowerCase() === status) && (!term || searchable.includes(term));
  });
  el("leadCount").textContent = `${filtered.length} de ${leads.length} solicitação(ões)`;
  renderLeads(filtered, plans);
}

function renderPlanRequestTable() {
  const requests = state.data?.plan_change_requests || [];
  const term = clean(el("planRequestSearch").value).toLowerCase();
  const status = clean(el("planRequestStatusFilter").value).toLowerCase();
  const filtered = requests.filter((request) => {
    const searchable = [
      request.company_name,
      request.current_plan_code,
      request.current_plan_name,
      request.requested_plan_code,
      request.requested_plan_name,
      request.message,
      request.requested_by_name,
    ].map(clean).join(" ").toLowerCase();
    return (!status || clean(request.status).toLowerCase() === status) && (!term || searchable.includes(term));
  });
  el("planRequestCount").textContent = `${filtered.length} de ${requests.length} solicitação(ões)`;
  renderPlanRequests(filtered);
}

function limitLabel(limit) {
  return limit === null || limit === undefined || limit === "" ? "sem limite" : `${number(limit)}`;
}

function renderPlanRequests(requests) {
  const rows = el("planRequestRows");
  rows.innerHTML = "";
  if (!requests.length) {
    rows.innerHTML = '<tr><td colspan="9">Nenhuma solicitação de plano encontrada.</td></tr>';
    return;
  }
  requests.forEach((request) => {
    const pending = clean(request.status).toLowerCase() === "pending";
    const usage = [
      `${number(request.vehicle_count)} veíc. / limite ${limitLabel(request.vehicle_limit_requested)}`,
      `${number(request.user_count)} usu. / limite ${limitLabel(request.user_limit_requested)}`,
    ].join("<br>");
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(dateTime(request.created_at))}</td>
      <td><strong>${escapeHtml(request.company_name || request.company_id)}</strong><br><small>${escapeHtml(request.company_document || "")}</small></td>
      <td>${statusBadge(request.request_type)}</td>
      <td>${escapeHtml(request.current_plan_name || request.current_plan_code || "-")}</td>
      <td><strong>${escapeHtml(request.requested_plan_name || request.requested_plan_code || "-")}</strong><br><small>${escapeHtml(money(request.requested_plan_price || 0))}</small></td>
      <td>${usage}</td>
      <td>${escapeHtml(request.message || "-")}</td>
      <td>${statusBadge(request.status)}</td>
      <td><div class="actions">
        <button type="button" data-approve-plan-request="${request.id}" ${pending ? "" : "disabled"}>Aprovar</button>
        <button type="button" class="danger" data-reject-plan-request="${request.id}" ${pending ? "" : "disabled"}>Recusar</button>
      </div></td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-approve-plan-request]").forEach((button) => {
    button.addEventListener("click", () => approvePlanRequest(Number(button.dataset.approvePlanRequest)));
  });
  rows.querySelectorAll("[data-reject-plan-request]").forEach((button) => {
    button.addEventListener("click", () => rejectPlanRequest(Number(button.dataset.rejectPlanRequest)));
  });
}

function renderLeads(leads, plans) {
  const rows = el("leadRows");
  rows.innerHTML = "";
  if (!leads.length) {
    rows.innerHTML = '<tr><td colspan="8">Nenhuma solicitação encontrada.</td></tr>';
    return;
  }
  const planOptions = (selected) => plans.map((plan) => `
    <option value="${escapeHtml(plan.code)}" ${plan.code === selected ? "selected" : ""}>${escapeHtml(plan.name || plan.code)}</option>
  `).join("");
  leads.forEach((lead) => {
    const pending = clean(lead.status).toLowerCase() === "novo";
    const autoActivated = clean(lead.status).toLowerCase() === "aprovado"
      && clean(lead.reviewed_by).toLowerCase() === "cadastro_publico";
    const actionHtml = pending
      ? `
        <button type="button" data-approve-lead="${lead.id}">Ativar demo</button>
        <button type="button" class="danger" data-reject-lead="${lead.id}">Recusar</button>
      `
      : `<span class="muted-action">${autoActivated ? "Autoativado" : "Sem ação"}</span>`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(dateTime(lead.created_at))}</td>
      <td><strong>${escapeHtml(lead.company)}</strong><br><small>${escapeHtml(lead.message || "")}</small></td>
      <td><div class="contact-lines"><strong>${escapeHtml(lead.name)}</strong><small>${escapeHtml(lead.email)}</small><small>${escapeHtml(lead.phone)}</small></div></td>
      <td>${escapeHtml(lead.document)}</td>
      <td><select class="inline-control" data-lead-plan="${lead.id}" ${pending ? "" : "disabled"}>${planOptions(lead.plan_code)}</select></td>
      <td><input class="trial-days" data-lead-days="${lead.id}" type="number" min="1" max="90" value="${escapeHtml(lead.trial_days || 30)}" ${pending ? "" : "disabled"}></td>
      <td>${statusBadge(lead.status)}</td>
      <td><div class="actions">
        ${actionHtml}
      </div></td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-approve-lead]").forEach((button) => {
    button.addEventListener("click", () => approveLead(Number(button.dataset.approveLead)));
  });
  rows.querySelectorAll("[data-reject-lead]").forEach((button) => {
    button.addEventListener("click", () => rejectLead(Number(button.dataset.rejectLead)));
  });
}

function renderCompanyTable() {
  const companies = state.data?.companies || [];
  const term = clean(el("companySearch").value).toLowerCase();
  const status = clean(el("companyStatusFilter").value).toLowerCase();
  const filtered = companies.filter((company) => {
    const companyStatus = clean(company.status).toLowerCase();
    const searchable = [
      company.id,
      company.name,
      company.nome,
      company.razao_social,
      company.plan_code,
      company.plano,
      company.status,
      statusLabel(companyStatus),
    ].map(clean).join(" ").toLowerCase();
    return (!status || companyStatus === status) && (!term || searchable.includes(term));
  });
  el("companyCount").textContent = `${filtered.length} de ${companies.length} empresa(s)`;
  renderCompanies(filtered);
}

function renderCompanies(companies) {
  const rows = el("companyRows");
  rows.innerHTML = "";
  if (!companies.length) {
    rows.innerHTML = '<tr><td colspan="6">Nenhuma empresa cadastrada.</td></tr>';
    return;
  }
  companies.forEach((company) => {
    const tr = document.createElement("tr");
    if (Number(company.id) === Number(state.selectedCompanyId)) tr.className = "selected";
    tr.innerHTML = `
      <td>${escapeHtml(company.id)}</td>
      <td>${escapeHtml(company.name || company.nome || company.razao_social || "-")}</td>
      <td>${statusBadge(company.status)}</td>
      <td>${escapeHtml(company.plan_code || company.plano || "-")}</td>
      <td>${escapeHtml(company.next_due_date || company.vencimento || "-")}</td>
      <td><div class="actions">
        <button type="button" class="secondary" data-select-company="${company.id}">Selecionar</button>
        <button type="button" data-company-status="active" data-company-id="${company.id}">Ativar</button>
        <button type="button" class="danger" data-company-status="suspended" data-company-id="${company.id}">Suspender</button>
      </div></td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-select-company]").forEach((button) => {
    button.addEventListener("click", async () => {
      state.selectedCompanyId = Number(button.dataset.selectCompany);
      await loadDashboard();
    });
  });
  rows.querySelectorAll("[data-company-status]").forEach((button) => {
    button.addEventListener("click", () => updateCompanyStatus(Number(button.dataset.companyId), button.dataset.companyStatus));
  });
}

function renderSelectedCompany(company, plans, usage, features) {
  el("selectedCompanyInfo").textContent = company
    ? `${companyDisplayName(company)} | ${statusLabel(company.status)}`
    : "Nenhuma empresa selecionada.";
  el("paymentCompanyInfo").textContent = company
    ? paymentCompanyInfoText(company)
    : "Escolha o cliente/empresa para gerar a cobrança.";
  renderPaymentCompanySelect(company);
  applyPaymentPlanDefaults(company);
  const select = el("planSelect");
  select.innerHTML = plans.map((plan) => `<option value="${escapeHtml(plan.code)}">${escapeHtml(plan.name || plan.code)}</option>`).join("");
  if (company?.plan_code) select.value = company.plan_code;
  el("usageGrid").innerHTML = Object.entries(usage || {}).map(([key, value]) => `
    <div class="info-item"><span>${escapeHtml(key)}</span><strong>${escapeHtml(number(value))}</strong></div>
  `).join("");
  const featureItems = Object.entries(features.features || features || {}).slice(0, 12);
  el("featuresGrid").innerHTML = featureItems.map(([key, value]) => `
    <div class="info-item"><span>${escapeHtml(key)}</span><strong>${escapeHtml(value ? "Ativo" : "Inativo")}</strong></div>
  `).join("");
}

function paymentCompanyInfoText(company) {
  const plan = planForCompany(company);
  const planName = plan?.name || company?.plan_code || company?.plano || "-";
  const planPrice = planMonthlyPrice(plan);
  const priceText = planPrice > 0 ? ` | Valor do plano ${money(planPrice)}` : "";
  return `Destinatário da cobrança: ${companyDisplayName(company)}${company.document ? ` | ${company.document}` : ""} | Plano ${planName}${priceText} | ${statusLabel(company.status)}. Tolerância: ${PAYMENT_GRACE_DAYS} dias; após isso, multa ${LATE_FINE_PERCENT}% + juros ${MONTHLY_INTEREST_PERCENT}% ao mês.`;
}

function renderPaymentCompanySelect(selectedCompany) {
  const select = el("paymentCompanySelect");
  const companies = state.data?.companies || [];
  select.innerHTML = companies.length
    ? companies.map((company) => `<option value="${escapeHtml(company.id)}">${escapeHtml(companyOptionLabel(company))}</option>`).join("")
    : '<option value="">Nenhuma empresa cadastrada</option>';
  if (selectedCompany?.id) select.value = String(selectedCompany.id);
}

function applyPaymentPlanDefaults(company, force = false) {
  const plan = planForCompany(company);
  const price = planMonthlyPrice(plan);
  const amountInput = el("paymentAmount");
  if (price > 0 && (force || !clean(amountInput.value))) {
    amountInput.value = formatMoneyInput(price);
  }
  if (company?.next_due_date && (force || !clean(el("paymentDue").value))) {
    el("paymentDue").value = clean(company.next_due_date).slice(0, 10);
  }
  updatePaymentLatePreview();
}

function updatePaymentLatePreview() {
  const baseAmount = parseMoney(el("paymentAmount").value);
  const dueDate = clean(el("paymentDue").value);
  if (!Number.isFinite(baseAmount) || baseAmount <= 0) {
    el("paymentLateAmount").value = "";
    return;
  }
  const adjusted = adjustedPaymentAmount(baseAmount, dueDate);
  const chargeableDays = daysAfterGrace(dueDate);
  el("paymentLateAmount").value = chargeableDays > 0
    ? `${formatMoneyInput(adjusted)} (${chargeableDays} dia(s) com juros)`
    : `${formatMoneyInput(adjusted)} (sem juros)`;
}

async function changePaymentCompany(event) {
  const companyId = Number(event.currentTarget.value || 0);
  if (!companyId || companyId === Number(state.selectedCompanyId)) return;
  state.selectedCompanyId = companyId;
  await loadDashboard();
  applyPaymentPlanDefaults(currentCompany(), true);
  switchTab("payments");
}

function renderPayments(payments) {
  const rows = el("paymentRows");
  rows.innerHTML = "";
  if (!payments.length) {
    rows.innerHTML = '<tr><td colspan="8">Nenhuma cobrança cadastrada.</td></tr>';
    return;
  }
  payments.forEach((payment) => {
    const paid = clean(payment.status).toLowerCase() === "paid";
    const boletoUrl = clean(payment.boleto_pdf_url);
    const boletoHtml = boletoUrl
      ? `<div class="boleto-cell"><strong>${escapeHtml(payment.boleto_our_number || payment.reference || "Gerado")}</strong><small>${escapeHtml(payment.boleto_digitable_line || "")}</small></div>`
      : '<span class="muted-action">Não gerado</span>';
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(payment.id)}</td>
      <td>${statusBadge(payment.status)}</td>
      <td>${escapeHtml(money(payment.amount || payment.valor || 0))}</td>
      <td>${escapeHtml(dateTime(payment.due_date))}</td>
      <td>${escapeHtml(dateTime(payment.paid_at))}</td>
      <td>${escapeHtml(payment.reference || "-")}</td>
      <td>${boletoHtml}</td>
      <td><div class="actions">
        <button type="button" data-generate-boleto="${payment.id}" ${paid ? "disabled" : ""}>${boletoUrl ? "Regerar boleto" : "Gerar boleto"}</button>
        <button type="button" class="secondary" data-open-boleto="${payment.id}" ${boletoUrl ? "" : "disabled"}>Abrir</button>
        <button type="button" class="secondary" data-copy-boleto="${payment.id}" ${boletoUrl ? "" : "disabled"}>Copiar envio</button>
        <button type="button" class="secondary" data-register-payment="${payment.id}" ${paid ? "disabled" : ""}>Registrar baixa</button>
      </div></td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-generate-boleto]").forEach((button) => {
    button.addEventListener("click", () => generateBoleto(Number(button.dataset.generateBoleto)));
  });
  rows.querySelectorAll("[data-open-boleto]").forEach((button) => {
    button.addEventListener("click", () => openBoleto(Number(button.dataset.openBoleto)));
  });
  rows.querySelectorAll("[data-copy-boleto]").forEach((button) => {
    button.addEventListener("click", () => copyBoletoMessage(Number(button.dataset.copyBoleto)));
  });
  rows.querySelectorAll("[data-register-payment]").forEach((button) => {
    button.addEventListener("click", () => registerPayment(Number(button.dataset.registerPayment)));
  });
}

function renderAudit(audit) {
  const rows = el("auditRows");
  rows.innerHTML = "";
  if (!audit.length) {
    rows.innerHTML = '<tr><td colspan="5">Sem eventos de auditoria.</td></tr>';
    return;
  }
  audit.forEach((log) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(log.action || "-")}</td>
      <td>${escapeHtml(log.company_id || "-")}</td>
      <td>${escapeHtml(log.actor || log.user || "-")}</td>
      <td>${escapeHtml(dateTime(log.created_at))}</td>
      <td><pre>${escapeHtml(JSON.stringify(log.metadata || {}, null, 2))}</pre></td>
    `;
    rows.appendChild(tr);
  });
}

async function updateCompanyStatus(companyId, status) {
  if (!window.confirm(`${status === "active" ? "Ativar" : "Suspender"} esta empresa?`)) return;
  try {
    await api(`/saas-admin/companies/${companyId}/status`, {method: "PUT", body: JSON.stringify({status})});
    toast("Status atualizado.");
    await loadDashboard();
  } catch (error) {
    toast(error.message, true);
  }
}

async function approveLead(leadId) {
  const lead = (state.data?.signup_leads || []).find((item) => Number(item.id) === Number(leadId)) || {};
  const planCode = clean(document.querySelector(`[data-lead-plan="${leadId}"]`)?.value);
  const trialDays = Number(document.querySelector(`[data-lead-days="${leadId}"]`)?.value || 30);
  if (!planCode || !Number.isInteger(trialDays) || trialDays < 1 || trialDays > 90) {
    return toast("Informe plano e quantidade de dias entre 1 e 90.", true);
  }
  if (!window.confirm(`Ativar demonstração de ${trialDays} dia(s) no plano ${planCode}?`)) return;
  try {
    const result = await api(`/saas-admin/signup-leads/${leadId}/approve-trial`, {
      method: "POST",
      body: JSON.stringify({plan_code: planCode, trial_days: trialDays}),
    });
    state.selectedCompanyId = Number(result.company_id || 0) || null;
    toast("Demonstração ativada.");
    openAccessModal({
      companyName: lead.company,
      username: result.temporary_username,
      password: result.temporary_password,
      trialDays,
      trialEnd: result.trial_end ? dateTime(result.trial_end) : "",
    });
    await loadDashboard();
    switchTab("dashboard");
  } catch (error) {
    toast(error.message, true);
  }
}

async function rejectLead(leadId) {
  const reason = window.prompt("Motivo da recusa", "");
  if (reason === null) return;
  try {
    await api(`/saas-admin/signup-leads/${leadId}/reject`, {
      method: "POST",
      body: JSON.stringify({reason: clean(reason) || null}),
    });
    toast("Solicitação recusada.");
    await loadDashboard();
  } catch (error) {
    toast(error.message, true);
  }
}

async function approvePlanRequest(requestId) {
  const request = (state.data?.plan_change_requests || []).find((item) => Number(item.id) === Number(requestId)) || {};
  const notes = window.prompt("Observação para aprovação", `Aprovado para ${request.requested_plan_name || request.requested_plan_code || "novo plano"}.`);
  if (notes === null) return;
  try {
    await api(`/saas-admin/plan-change-requests/${requestId}/approve`, {
      method: "POST",
      body: JSON.stringify({notes: clean(notes) || null}),
    });
    state.selectedCompanyId = Number(request.company_id || state.selectedCompanyId || 0) || null;
    toast("Solicitação de plano aprovada.");
    await loadDashboard();
    switchTab("planRequests");
  } catch (error) {
    toast(error.message, true);
  }
}

async function rejectPlanRequest(requestId) {
  const notes = window.prompt("Motivo da recusa", "");
  if (notes === null) return;
  try {
    await api(`/saas-admin/plan-change-requests/${requestId}/reject`, {
      method: "POST",
      body: JSON.stringify({notes: clean(notes) || null}),
    });
    toast("Solicitação de plano recusada.");
    await loadDashboard();
    switchTab("planRequests");
  } catch (error) {
    toast(error.message, true);
  }
}

async function changePlan(event) {
  event.preventDefault();
  if (!state.selectedCompanyId) return toast("Selecione uma empresa.", true);
  try {
    await api(`/saas-admin/companies/${state.selectedCompanyId}/plan`, {
      method: "PUT",
      body: JSON.stringify({plan_code: clean(el("planSelect").value), reason: clean(el("planReason").value)}),
    });
    el("planReason").value = "";
    toast("Plano atualizado.");
    await loadDashboard();
  } catch (error) {
    toast(error.message, true);
  }
}

async function resetAdminAccess() {
  if (!state.selectedCompanyId) return toast("Selecione uma empresa.", true);
  if (!window.confirm("Gerar uma nova senha temporária para o administrador desta empresa?")) return;
  try {
    const result = await api(`/saas-admin/companies/${state.selectedCompanyId}/admin-access`, {method: "POST"});
    const company = currentCompany() || {};
    openAccessModal({
      companyName: company.name || company.nome || company.razao_social || "",
      username: result.username,
      password: result.temporary_password,
    });
    toast("Acesso administrativo gerado.");
  } catch (error) {
    toast(error.message, true);
  }
}

async function createPayment(event) {
  event.preventDefault();
  const companyId = Number(el("paymentCompanySelect").value || state.selectedCompanyId || 0);
  if (!companyId) return toast("Selecione o cliente da cobrança.", true);
  state.selectedCompanyId = companyId;
  const amount = parseMoney(el("paymentAmount").value);
  if (!Number.isFinite(amount) || amount <= 0) return toast("Informe um valor de cobrança válido.", true);
  const adjustedAmount = adjustedPaymentAmount(amount, clean(el("paymentDue").value));
  const shouldGenerateBoleto = Boolean(el("paymentGenerateBoleto")?.checked);
  try {
    const created = await api("/saas-admin/payments", {
      method: "POST",
      body: JSON.stringify({
        company_id: state.selectedCompanyId,
        amount: adjustedAmount,
        due_date: clean(el("paymentDue").value) || null,
        notes: clean(el("paymentNotes").value) || null,
      }),
    });
    event.currentTarget.reset();
    const paymentId = Number((created.payment || {}).id || 0);
    if (shouldGenerateBoleto && paymentId) {
      const boleto = await api(`/saas-admin/payments/${paymentId}/gerar-boleto`, {method: "POST"});
      const payment = boleto.payment || created.payment || {};
      toast("Cobrança e boleto criados.");
      if (payment.boleto_pdf_url) {
        await copyText(boletoMessage(payment), "Mensagem de envio copiada.");
      }
    } else {
      toast("Cobrança criada.");
    }
    await loadDashboard();
    switchTab("payments");
  } catch (error) {
    toast(error.message, true);
  }
}

function findPayment(paymentId) {
  return (state.data?.payments || []).find((item) => Number(item.id) === Number(paymentId)) || null;
}

async function generateBoleto(paymentId) {
  try {
    const result = await api(`/saas-admin/payments/${paymentId}/gerar-boleto`, {method: "POST"});
    toast("Boleto gerado.");
    await loadDashboard();
    const payment = result.payment || findPayment(paymentId);
    if (payment?.boleto_pdf_url) {
      await copyText(boletoMessage(payment), "Mensagem de envio copiada.");
    }
  } catch (error) {
    toast(error.message, true);
  }
}

function openBoleto(paymentId) {
  const payment = findPayment(paymentId);
  const url = absoluteUrl(payment?.boleto_pdf_url);
  if (!url) return toast("Boleto ainda nao gerado.", true);
  window.open(url, "_blank", "noopener");
}

async function copyBoletoMessage(paymentId) {
  const payment = findPayment(paymentId);
  if (!payment?.boleto_pdf_url) return toast("Boleto ainda nao gerado.", true);
  await copyText(boletoMessage(payment), "Mensagem de envio copiada.");
}

async function registerPayment(paymentId) {
  const reference = window.prompt("Referência do pagamento", "");
  if (reference === null) return;
  try {
    await api(`/saas-admin/payments/${paymentId}/registrar-pagamento`, {
      method: "POST",
      body: JSON.stringify({method: "manual", reference: clean(reference) || null}),
    });
    toast("Pagamento registrado.");
    await loadDashboard();
  } catch (error) {
    toast(error.message, true);
  }
}

async function runOverdue() {
  const graceDays = Math.max(parseInt(clean(el("overdueDays").value) || "0", 10) || 0, 0);
  if (!window.confirm(`Verificar inadimplência com ${graceDays} dia(s) de tolerância?`)) return;
  try {
    const result = await api("/saas-admin/billing/run-overdue-check", {
      method: "POST",
      body: JSON.stringify({grace_days: graceDays}),
    });
    toast(`Assinaturas suspensas: ${(result.summary || {}).suspended || 0}`);
    await loadDashboard();
  } catch (error) {
    toast(error.message, true);
  }
}

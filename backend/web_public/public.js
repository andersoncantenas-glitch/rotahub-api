const API_BASE = "/api/v1";

const FALLBACK_PLANS = [
  {
    code: "starter",
    name: "Inicial 5 Veículos",
    monthly_price: 199,
    vehicle_limit: 5,
    user_limit: 6,
    description: "Para começar organizado: planejamento, recebimentos, custos e app do motorista em até 5 veículos.",
    summary: "O básico bem feito para parar de controlar rota no improviso.",
    features: ["Cadastros da operação", "Planejamento de rotas", "Recebimentos e custos", "App do motorista"],
  },
  {
    code: "growth",
    name: "Crescimento 10 Veículos",
    monthly_price: 399,
    vehicle_limit: 10,
    user_limit: 15,
    description: "Para controlar ocorrências operacionais, rotas, financeiro básico e análise de custos em até 10 veículos.",
    summary: "Mais visão para saber onde a rota está dando lucro ou prejuízo.",
    features: ["Tudo do Inicial", "Ocorrências operacionais", "Rotas", "Análise de custos", "Financeiro básico"],
  },
  {
    code: "professional",
    name: "Profissional 15 Veículos",
    monthly_price: 699,
    vehicle_limit: 15,
    user_limit: 30,
    description: "Para operações mais exigentes: escala, relatórios avançados, rotas e controles completos em até 15 veículos.",
    summary: "Controle completo para gestor acompanhar equipe, frota e resultado.",
    features: ["Tudo do Crescimento", "Escala", "Relatórios avançados", "Controle por perfis", "Gestão completa"],
  },
  {
    code: "enterprise",
    name: "Empresarial Mais Veículos",
    monthly_price: 999,
    vehicle_limit: null,
    user_limit: null,
    price_label: "Sob consulta",
    description: "Para empresas com mais de 15 veículos, API, suporte prioritário e contrato ajustado à operação.",
    summary: "Plano sob medida para operação maior, com implantação acompanhada.",
    features: ["Tudo do Profissional", "Mais de 15 veículos", "API", "Contrato customizado", "Suporte prioritário"],
  },
];

function el(id) {
  return document.getElementById(id);
}

function clean(value) {
  return String(value || "").trim();
}

function digits(value) {
  return clean(value).replace(/\D+/g, "");
}

function money(value) {
  const number = Number(value || 0);
  return number.toLocaleString("pt-BR", {style: "currency", currency: "BRL"});
}

function planPrice(plan) {
  return clean(plan.price_label) || money(plan.monthly_price || 0);
}

function planLimitLabel(plan) {
  if (plan.vehicle_limit === null || plan.vehicle_limit === undefined) return "Frota sob contrato";
  return `Até ${plan.vehicle_limit} veículos`;
}

function transitionToLogin(url = "/app/index.html") {
  const transition = el("loginTransition");
  if (transition.classList.contains("is-visible")) return;
  transition.classList.remove("hidden");
  requestAnimationFrame(() => transition.classList.add("is-visible"));
  window.setTimeout(() => {
    window.location.href = url;
  }, 1500);
}

function escapeHtml(value) {
  return clean(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function apiRequest(path, options = {}) {
  const headers = new Headers(options.headers || {});
  if (options.body && !(options.body instanceof FormData)) {
    headers.set("Content-Type", "application/json");
  }
  const response = await fetch(`${API_BASE}${path}`, {...options, headers});
  const text = await response.text();
  const data = text ? JSON.parse(text) : null;
  if (!response.ok) {
    const detail = data && data.detail ? data.detail : "Não foi possível concluir a solicitação.";
    throw new Error(Array.isArray(detail) ? detail.map((item) => item.msg).join("; ") : detail);
  }
  return data;
}

function renderPlans(plans) {
  const safePlans = Array.isArray(plans) && plans.length ? plans : FALLBACK_PLANS;
  el("plansGrid").innerHTML = safePlans.map((plan, index) => {
    const features = Array.isArray(plan.features) && plan.features.length ? plan.features : FALLBACK_PLANS[index % FALLBACK_PLANS.length].features;
    const isFeatured = plan.code === "professional" || index === 2;
    const nextPlan = safePlans[index + 1];
    return `
      <article class="plan-card ${isFeatured ? "featured" : ""}">
        <div class="plan-topline">
          <span>${escapeHtml(planLimitLabel(plan))}</span>
          ${isFeatured ? "<em>Recomendado</em>" : ""}
        </div>
        <div class="plan-copy">
          <h3>${escapeHtml(plan.name || plan.code)}</h3>
          <strong>${escapeHtml(plan.summary || "Plano RotaHub para gestão operacional.")}</strong>
          <p>${escapeHtml(plan.description || "Módulo RotaHub para gestão operacional.")}</p>
        </div>
        <div class="price">
          <strong>${escapeHtml(planPrice(plan))}</strong>
          <span>${clean(plan.price_label) ? "" : "/mês"}</span>
        </div>
        <ul>${features.slice(0, 6).map((feature) => `<li>${escapeHtml(feature)}</li>`).join("")}</ul>
        ${nextPlan ? `<p class="plan-upgrade">Quer ampliar depois? Conheça também o <strong>${escapeHtml(nextPlan.name || nextPlan.code)}</strong>.</p>` : '<p class="plan-upgrade">Implantação acompanhada para operações maiores.</p>'}
        <a class="primary-button" href="#cadastro" data-plan="${escapeHtml(plan.code)}">Escolher apos o teste</a>
      </article>
    `;
  }).join("");

  const select = el("signupPlan");
  select.innerHTML = safePlans.map((plan) => `<option value="${escapeHtml(plan.code)}">${escapeHtml(plan.name || plan.code)} - ${escapeHtml(planPrice(plan))}</option>`).join("");
  document.querySelectorAll("[data-plan]").forEach((button) => {
    button.addEventListener("click", () => {
      select.value = button.getAttribute("data-plan") || select.value;
    });
  });
}

async function loadPlans() {
  try {
    const data = await apiRequest("/public/plans");
    renderPlans(data.plans);
  } catch (_error) {
    renderPlans(FALLBACK_PLANS);
  }
}

async function onSignup(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const status = el("signupStatus");
  status.className = "form-status";
  status.textContent = "";
  const data = Object.fromEntries(new FormData(form).entries());
  data.document = digits(data.document);
  if (![11, 14].includes(data.document.length)) {
    status.className = "form-status error";
    status.textContent = "Informe um CPF ou CNPJ válido.";
    return;
  }
  if (data.password !== data.password_confirm) {
    status.className = "form-status error";
    status.textContent = "A confirmacao da senha nao confere.";
    return;
  }
  delete data.password_confirm;
  try {
    el("signupButton").disabled = true;
    const result = await apiRequest("/public/signup", {method: "POST", body: JSON.stringify(data)});
    status.textContent = result.message || "Cadastro recebido.";
    el("signupSuccessLogin").href = result.next_url || "/app/index.html";
    el("signupSuccessModal").classList.remove("hidden");
    form.reset();
  } catch (error) {
    status.className = "form-status error";
    status.textContent = error.message;
  } finally {
    el("signupButton").disabled = false;
  }
}

function onClientAccess(event) {
  event.preventDefault();
  const status = el("clientAccessStatus");
  const documentValue = digits(el("clientDocument").value);
  status.className = "";
  if (![11, 14].includes(documentValue.length)) {
    status.className = "error";
    status.textContent = "Informe CPF ou CNPJ válido.";
    return;
  }
  localStorage.setItem("rotahub_client_document", documentValue);
  status.textContent = "Redirecionando para o login...";
  transitionToLogin("/app/index.html");
}

document.addEventListener("DOMContentLoaded", () => {
  loadPlans();
  el("signupForm").addEventListener("submit", onSignup);
  el("clientAccessForm").addEventListener("submit", onClientAccess);
  el("signupSuccessLogin").addEventListener("click", (event) => {
    event.preventDefault();
    transitionToLogin(event.currentTarget.href);
  });
});

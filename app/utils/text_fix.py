import re


_MOJIBAKE_MARKERS = (
    "\u00c3",
    "\u00c2",
    "\u00e2",
    "\u0192",
    "\u00c6",
    "\u00a2",
    "\u20ac",
    "\u2122",
    "\u0153",
    "\u017e",
    "\ufffd",
    "\u0102",
    "\u0103",
    "\u0161",
    "\u2030",
    "\x89",
    "\x9a",
)

_UI_TERMS = {
    "acao": "ação",
    "acoes": "ações",
    "administracao": "administração",
    "alteracao": "alteração",
    "alteracoes": "alterações",
    "antencao": "atenção",
    "aplicacao": "aplicação",
    "aplicacoes": "aplicações",
    "area": "área",
    "areas": "áreas",
    "atencao": "atenção",
    "atualizacao": "atualização",
    "atualizacoes": "atualizações",
    "automatica": "automática",
    "automaticas": "automáticas",
    "automatico": "automático",
    "automaticos": "automáticos",
    "autenticacao": "autenticação",
    "cabecalho": "cabeçalho",
    "cabecalhos": "cabeçalhos",
    "cedula": "cédula",
    "cedulas": "cédulas",
    "codigo": "código",
    "codigos": "códigos",
    "comparacao": "comparação",
    "comunicacao": "comunicação",
    "conexao": "conexão",
    "configuracao": "configuração",
    "configuracoes": "configurações",
    "confirmacao": "confirmação",
    "conferencia": "conferência",
    "conferencias": "conferências",
    "conclusao": "conclusão",
    "conteudo": "conteúdo",
    "conteudos": "conteúdos",
    "conversao": "conversão",
    "correcao": "correção",
    "correcoes": "correções",
    "criacao": "criação",
    "criterio": "critério",
    "criterios": "critérios",
    "descricao": "descrição",
    "descricoes": "descrições",
    "deteccao": "detecção",
    "diaria": "diária",
    "diarias": "diárias",
    "distribuicao": "distribuição",
    "edicao": "edição",
    "eletronico": "eletrônico",
    "eletronicos": "eletrônicos",
    "emissao": "emissão",
    "endereco": "endereço",
    "enderecos": "endereços",
    "especificacao": "especificação",
    "exportacao": "exportação",
    "exclusao": "exclusão",
    "execucao": "execução",
    "finalizacao": "finalização",
    "formulario": "formulário",
    "formularios": "formulários",
    "funcao": "função",
    "funcoes": "funções",
    "geracao": "geração",
    "grafico": "gráfico",
    "graficos": "gráficos",
    "historico": "histórico",
    "icone": "ícone",
    "icones": "ícones",
    "impressao": "impressão",
    "importacao": "importação",
    "importacoes": "importações",
    "inicio": "início",
    "informacao": "informação",
    "informacoes": "informações",
    "integracao": "integração",
    "lancada": "lançada",
    "lancadas": "lançadas",
    "lancado": "lançado",
    "lancados": "lançados",
    "lancamento": "lançamento",
    "lancamentos": "lançamentos",
    "localizacao": "localização",
    "logistica": "logística",
    "manutencao": "manutenção",
    "media": "média",
    "medias": "médias",
    "metodo": "método",
    "metodos": "métodos",
    "modulo": "módulo",
    "modulos": "módulos",
    "numero": "número",
    "observacao": "observação",
    "observacoes": "observações",
    "operacao": "operação",
    "operacoes": "operações",
    "opcao": "opção",
    "opcoes": "opções",
    "pagina": "página",
    "paginas": "páginas",
    "padrao": "padrão",
    "padroes": "padrões",
    "parametro": "parâmetro",
    "parametros": "parâmetros",
    "pendencia": "pendência",
    "pendencias": "pendências",
    "permissao": "permissão",
    "permissoes": "permissões",
    "periodo": "período",
    "politica": "política",
    "preco": "preço",
    "precos": "preços",
    "prestacao": "prestação",
    "prestacoes": "prestações",
    "programacao": "programação",
    "programacoes": "programações",
    "relatorio": "relatório",
    "relatorios": "relatórios",
    "restricao": "restrição",
    "restricoes": "restrições",
    "revisao": "revisão",
    "selecao": "seleção",
    "selecoes": "seleções",
    "sertao": "sertão",
    "servico": "serviço",
    "servicos": "serviços",
    "sessao": "sessão",
    "saida": "saída",
    "saidas": "saídas",
    "situacao": "situação",
    "sincronizacao": "sincronização",
    "solucao": "solução",
    "substituicao": "substituição",
    "sumario": "sumário",
    "tecnico": "técnico",
    "tecnicos": "técnicos",
    "transacao": "transação",
    "transacoes": "transações",
    "ultima": "última",
    "ultimas": "últimas",
    "ultimo": "último",
    "ultimos": "últimos",
    "usuario": "usuário",
    "usuarios": "usuários",
    "util": "útil",
    "utilitario": "utilitário",
    "utilitarios": "utilitários",
    "validacao": "validação",
    "validacoes": "validações",
    "variacao": "variação",
    "veiculo": "veículo",
    "veiculos": "veículos",
    "versao": "versão",
    "versoes": "versões",
    "visualizacao": "visualização",
    "visualizacoes": "visualizações",
    "vinculacao": "vinculação",
}

_MESSAGE_TERMS = {
    "ate": "até",
    "critico": "crítico",
    "critica": "crítica",
    "disponivel": "disponível",
    "disponiveis": "disponíveis",
    "excluida": "excluída",
    "excluido": "excluído",
    "indisponivel": "indisponível",
    "indisponiveis": "indisponíveis",
    "invalida": "inválida",
    "invalidas": "inválidas",
    "invalido": "inválido",
    "invalidos": "inválidos",
    "ja": "já",
    "maxima": "máxima",
    "maximas": "máximas",
    "maximo": "máximo",
    "maximos": "máximos",
    "minima": "mínima",
    "minimas": "mínimas",
    "minimo": "mínimo",
    "minimos": "mínimos",
    "nao": "não",
    "obrigatoria": "obrigatória",
    "obrigatorias": "obrigatórias",
    "obrigatorio": "obrigatório",
    "obrigatorios": "obrigatórios",
    "possivel": "possível",
    "possiveis": "possíveis",
    "reabertura": "reabertura",
    "rapida": "rápida",
    "rapidas": "rápidas",
    "rapido": "rápido",
    "rapidos": "rápidos",
    "sera": "será",
    "suite": "suíte",
}

_PHRASE_REPLACEMENTS = (
    ("pre visualizacao", "pré-visualização"),
    ("pre-visualizacao", "pré-visualização"),
    ("nao e possivel", "não é possível"),
    ("ja esta", "já está"),
    ("ultima atualizacao", "última atualização"),
    ("ultima versao", "última versão"),
    ("versao local", "versão local"),
    ("versao disponivel", "versão disponível"),
)


def _mojibake_score(text: str) -> int:
    if not text:
        return 0
    return sum(text.count(ch) for ch in _MOJIBAKE_MARKERS)


def _preserve_case(source: str, target: str) -> str:
    if not source:
        return target
    if source.upper() == source:
        return target.upper()
    if source.lower() == source:
        return target.lower()
    if source.istitle():
        return " ".join(part.capitalize() for part in target.split(" "))
    if source[:1].isupper():
        return target[:1].upper() + target[1:]
    return target


def _replace_terms(text: str, mapping: dict[str, str]) -> str:
    if not text:
        return text
    for raw, cooked in mapping.items():
        pattern = re.compile(rf"(?<!\w){re.escape(raw)}(?!\w)", re.IGNORECASE)
        text = pattern.sub(lambda m: _preserve_case(m.group(0), cooked), text)
    return text


def _replace_phrases(text: str) -> str:
    if not text:
        return text
    for raw, cooked in _PHRASE_REPLACEMENTS:
        pattern = re.compile(re.escape(raw), re.IGNORECASE)
        text = pattern.sub(lambda m: _preserve_case(m.group(0), cooked), text)
    return text


def fix_mojibake_text(value):
    if not isinstance(value, str) or not value:
        return value

    best = value
    best_score = _mojibake_score(best)
    if best_score == 0:
        return value

    for _ in range(4):
        improved = False
        # latin-1/cp1252 cobrem o mojibake classico (ex.: "Programa\u00c3\u00a7\u00c3\u00a3o")
        # cp1250/iso8859_2 cobrem variacoes do tipo "\u0102\u0161" / "\u0102\u2030".
        for enc in ("latin-1", "cp1252", "cp1250", "iso8859_2"):
            try:
                candidate = best.encode(enc).decode("utf-8")
            except Exception:
                continue
            cand_score = _mojibake_score(candidate)
            if cand_score < best_score:
                best, best_score = candidate, cand_score
                improved = True
        if not improved:
            break

    # fallback leve para sequencias restantes comuns
    best = best.replace("\u00e2\u20ac\u201d", "-").replace("\u00e2\u20ac\u201c", "-")
    best = best.replace("\u00e2\u20ac\u02dc", "\u2191").replace("\u00e2\u20ac\u0153", "\u2193")
    return best


def normalize_ui_text(value, aggressive: bool = True):
    if not isinstance(value, str) or not value:
        return value

    text = fix_mojibake_text(value)
    text = _replace_phrases(text)
    text = _replace_terms(text, _UI_TERMS)
    if aggressive:
        text = _replace_terms(text, _MESSAGE_TERMS)
    return text


def normalize_ui_collection(values, aggressive: bool = False):
    if isinstance(values, tuple):
        return tuple(normalize_ui_text(v, aggressive=aggressive) for v in values)
    if isinstance(values, list):
        return [normalize_ui_text(v, aggressive=aggressive) for v in values]
    return values

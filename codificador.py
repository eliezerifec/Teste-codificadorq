"""
codificador.py — Motor de codificacao com IA (OpenAI GPT)
Exporta: CodificadorIA, TIPOS_PERGUNTA, MODOS_RESPOSTA
"""

import os
import re
import time
from typing import Callable

# ---------------------------------------------------------------------------
# TAXONOMIAS  (usadas pelo Streamlit e pelo Tkinter)
# ---------------------------------------------------------------------------

TIPOS_PERGUNTA: dict[str, dict] = {
    "satisfacao": {
        "label": "Satisfacao / Motivo",
        "instrucao": (
            "Classifique a resposta em uma categoria que capture o MOTIVO ou "
            "SENTIMENTO principal expresso. Use substantivos curtos (2-4 palavras). "
            "Exemplos de categorias: 'Atendimento', 'Preco', 'Qualidade do produto', "
            "'Praticidade', 'Localizacao'."
        ),
    },
    "reconhecimento_marca": {
        "label": "Reconhecimento de Marca",
        "instrucao": (
            "Extraia o NOME DA MARCA ou PRODUTO mencionado na resposta, "
            "sem alterar a grafia. Se houver mais de uma marca, liste todas "
            "separadas por ';'. Se nao houver marca clara, use 'Nao identificado'."
        ),
    },
    "definicao_palavra": {
        "label": "Definicao em uma Palavra",
        "instrucao": (
            "O respondente deveria associar o tema a UMA PALAVRA. "
            "Extraia essa palavra exatamente como escrita, normalizando apenas "
            "capitalizacao. Se houver mais de uma palavra, escolha a mais relevante."
        ),
    },
    "local_moradia": {
        "label": "Local de Moradia",
        "instrucao": (
            "Extraia o BAIRRO, CIDADE ou REGIAO mencionado. "
            "Normalize abreviacoes comuns (ex: 'Cpa' -> 'Copacabana'). "
            "Se nao for possivel identificar, use 'Nao informado'."
        ),
    },
    "livre": {
        "label": "Resposta Livre / Personalizado",
        "instrucao": (
            "Use o contexto fornecido para criar ou imputar categorias "
            "que melhor representem o conteudo da resposta."
        ),
    },
}

MODOS_RESPOSTA: dict[str, dict] = {
    "simples": {
        "label": "Simples (uma categoria)",
        "instrucao": "Retorne APENAS UMA categoria para esta resposta.",
    },
    "multipla": {
        "label": "Multipla (varias categorias)",
        "instrucao": (
            "Retorne TODAS as categorias aplicaveis, separadas por ';'. "
            "Nao repita categorias."
        ),
    },
    "semiaberta_simples": {
        "label": "Semiaberta Simples",
        "instrucao": (
            "Se a resposta se encaixar em uma das categorias pre-definidas, "
            "use-a exatamente. Caso contrario, crie UMA nova categoria curta "
            "que represente o conteudo. Retorne apenas uma categoria."
        ),
    },
    "semiaberta_multipla": {
        "label": "Semiaberta Multipla",
        "instrucao": (
            "Para cada ideia presente na resposta: se encaixar em uma categoria "
            "pre-definida, use-a; se nao, crie uma nova categoria curta. "
            "Retorne todas separadas por ';'. Nao repita categorias."
        ),
    },
}


# ---------------------------------------------------------------------------
# CLASSE PRINCIPAL
# ---------------------------------------------------------------------------

class CodificadorIA:
    """
    Codifica respostas abertas usando OpenAI GPT.

    Uso:
        cod = CodificadorIA()
        resultado = cod.codificar_lote_modo(
            respostas,
            tipo="satisfacao",
            modo="simples",
            contexto_custom="",
            categorias_imputacao=[],
            categorias_anteriores=[],
            callback_progresso=None,
        )
        # resultado["resultado"]  -> list[str]  (modo simples/multipla)
        # resultado["imputado"]   -> list[str]  (modo semiaberto)
        # resultado["novo"]       -> list[str]  (modo semiaberto)
    """

    MODELO = "gpt-4o-mini"
    MAX_TENTATIVAS = 3
    PAUSA_ENTRE_CHAMADAS = 0.15  # segundos

    def __init__(self):
        self.categorias: list[str] = []
        self._client = None

    # ------------------------------------------------------------------
    # Cliente OpenAI (lazy init)
    # ------------------------------------------------------------------

    def _get_client(self):
        if self._client is None:
            try:
                from openai import OpenAI
            except ImportError as exc:
                raise ImportError(
                    "Pacote 'openai' nao encontrado. "
                    "Execute: pip install openai"
                ) from exc
            api_key = os.getenv("OPENAI_API_KEY", "")
            if not api_key:
                raise EnvironmentError(
                    "OPENAI_API_KEY nao configurada. "
                    "Defina a variavel de ambiente ou configure em Secrets."
                )
            self._client = OpenAI(api_key=api_key)
        return self._client

    # ------------------------------------------------------------------
    # Metodo publico principal
    # ------------------------------------------------------------------

    def codificar_lote_modo(
        self,
        respostas: list[str],
        tipo: str = "livre",
        modo: str = "simples",
        contexto_custom: str = "",
        categorias_imputacao: list[str] | None = None,
        categorias_anteriores: list[str] | None = None,
        callback_progresso: Callable | None = None,
    ) -> dict[str, list[str]]:
        """
        Codifica uma lista de respostas.

        Retorna dict com chaves:
          - "resultado"  (modo simples ou multipla)
          - "imputado" + "novo"  (modo semiaberto)
        """
        cats_imp = categorias_imputacao or []
        cats_ant = categorias_anteriores or []
        total = len(respostas)
        eh_semi = "semi" in modo

        resultados_imp: list[str] = []
        resultados_novo: list[str] = []
        resultados_simples: list[str] = []

        for i, resposta in enumerate(respostas):
            texto = str(resposta).strip()
            if not texto or texto.lower() in ("nan", "none", ""):
                if eh_semi:
                    resultados_imp.append("SEM_RESPOSTA")
                    resultados_novo.append("")
                else:
                    resultados_simples.append("SEM_RESPOSTA")
                if callback_progresso:
                    callback_progresso(i, total, texto, "SEM_RESPOSTA")
                continue

            try:
                if eh_semi:
                    imp, novo = self._codificar_semiaberto(
                        texto, tipo, modo, contexto_custom,
                        cats_imp, cats_ant,
                    )
                    resultados_imp.append(imp)
                    resultados_novo.append(novo)
                    categoria_log = f"imp={imp} novo={novo}"
                else:
                    cat = self._codificar_simples(
                        texto, tipo, modo, contexto_custom,
                        cats_imp, cats_ant,
                    )
                    resultados_simples.append(cat)
                    categoria_log = cat

            except Exception as exc:
                categoria_log = f"ERRO: {exc}"
                if eh_semi:
                    resultados_imp.append("ERRO")
                    resultados_novo.append("")
                else:
                    resultados_simples.append("ERRO")

            if callback_progresso:
                callback_progresso(i, total, texto, categoria_log)

            time.sleep(self.PAUSA_ENTRE_CHAMADAS)

        if eh_semi:
            return {"imputado": resultados_imp, "novo": resultados_novo}
        return {"resultado": resultados_simples}

    # ------------------------------------------------------------------
    # Codificacao simples / multipla
    # ------------------------------------------------------------------

    def _codificar_simples(
        self,
        resposta: str,
        tipo: str,
        modo: str,
        contexto_custom: str,
        cats_imp: list[str],
        cats_ant: list[str],
    ) -> str:
        system_prompt = self._montar_system(tipo, modo, contexto_custom, cats_ant)
        user_prompt = self._montar_user(resposta, cats_imp)

        resposta_ia = self._chamar_api(system_prompt, user_prompt)
        return self._limpar_resposta(resposta_ia)

    # ------------------------------------------------------------------
    # Codificacao semiaberta
    # ------------------------------------------------------------------

    def _codificar_semiaberto(
        self,
        resposta: str,
        tipo: str,
        modo: str,
        contexto_custom: str,
        cats_imp: list[str],
        cats_ant: list[str],
    ) -> tuple[str, str]:
        """
        Retorna (categorias_imputadas, categorias_novas).
        Cada valor pode ser string simples ou multipla (separada por ;).
        """
        system_prompt = self._montar_system_semi(tipo, modo, contexto_custom, cats_imp, cats_ant)
        user_prompt = self._montar_user(resposta, cats_imp)

        resposta_ia = self._chamar_api(system_prompt, user_prompt)
        return self._parsear_semi(resposta_ia, cats_imp)

    # ------------------------------------------------------------------
    # Construtores de prompt
    # ------------------------------------------------------------------

    def _montar_system(
        self,
        tipo: str,
        modo: str,
        contexto_custom: str,
        cats_ant: list[str],
    ) -> str:
        info_tipo = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        info_modo = MODOS_RESPOSTA.get(modo, MODOS_RESPOSTA["simples"])

        partes = [
            "Voce e um codificador de pesquisas qualitativas experiente.",
            "",
            f"TIPO DE PERGUNTA: {info_tipo['label']}",
            f"Instrucao de tipo: {info_tipo['instrucao']}",
            "",
            f"MODO DE RESPOSTA: {info_modo['label']}",
            f"Instrucao de modo: {info_modo['instrucao']}",
        ]

        if contexto_custom.strip():
            partes += ["", f"CONTEXTO DA PESQUISA: {contexto_custom.strip()}"]

        if cats_ant:
            partes += [
                "",
                "CATEGORIAS JA CRIADAS EM PESQUISAS ANTERIORES (prefira reutilizar):",
                ", ".join(cats_ant),
            ]

        partes += [
            "",
            "REGRAS GERAIS:",
            "- Retorne SOMENTE a(s) categoria(s), sem explicacoes, sem aspas, sem pontuacao final.",
            "- Use letras maiusculas na primeira palavra apenas (Title case so para nomes proprios).",
            "- Seja conciso: categorias com 1-4 palavras.",
            "- Nunca retorne a resposta original como categoria.",
        ]

        return "\n".join(partes)

    def _montar_system_semi(
        self,
        tipo: str,
        modo: str,
        contexto_custom: str,
        cats_imp: list[str],
        cats_ant: list[str],
    ) -> str:
        info_tipo = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        info_modo = MODOS_RESPOSTA.get(modo, MODOS_RESPOSTA["semiaberta_simples"])
        eh_multi = "multipla" in modo

        partes = [
            "Voce e um codificador de pesquisas qualitativas experiente.",
            "",
            f"TIPO DE PERGUNTA: {info_tipo['label']}",
            f"Instrucao de tipo: {info_tipo['instrucao']}",
            "",
            f"MODO: {info_modo['label']}",
        ]

        if cats_imp:
            partes += [
                "",
                "CATEGORIAS PRE-DEFINIDAS (use sempre que possivel):",
                "\n".join(f"  - {c}" for c in cats_imp),
            ]

        if cats_ant:
            partes += [
                "",
                "CATEGORIAS DE PESQUISAS ANTERIORES (referencia adicional):",
                ", ".join(cats_ant),
            ]

        if contexto_custom.strip():
            partes += ["", f"CONTEXTO: {contexto_custom.strip()}"]

        sep = ";" if eh_multi else ""
        partes += [
            "",
            "FORMATO DE SAIDA OBRIGATORIO — retorne exatamente duas linhas:",
            "IMPUTADO: <categoria(s) das pre-definidas" + (", separadas por ;" if eh_multi else "") + " ou SEM_IMPUTACAO>",
            "NOVO: <categoria(s) novas nao cobertas pelas pre-definidas" + (", separadas por ;" if eh_multi else "") + " ou vazio>",
            "",
            "Nao adicione mais nada alem dessas duas linhas.",
            "Se a resposta se encaixar em uma pre-definida, coloque-a em IMPUTADO e deixe NOVO vazio.",
            "Se nao se encaixar em nenhuma pre-definida, deixe IMPUTADO como SEM_IMPUTACAO e crie uma categoria em NOVO.",
        ]

        return "\n".join(partes)

    def _montar_user(self, resposta: str, cats_imp: list[str]) -> str:
        return f'Resposta: "{resposta}"'

    # ------------------------------------------------------------------
    # Chamada a API
    # ------------------------------------------------------------------

    def _chamar_api(self, system: str, user: str) -> str:
        client = self._get_client()
        ultima_excecao = None

        for tentativa in range(self.MAX_TENTATIVAS):
            try:
                completion = client.chat.completions.create(
                    model=self.MODELO,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                    temperature=0.0,
                    max_tokens=120,
                )
                return completion.choices[0].message.content.strip()
            except Exception as exc:
                ultima_excecao = exc
                espera = 2 ** tentativa
                time.sleep(espera)

        raise RuntimeError(
            f"API falhou apos {self.MAX_TENTATIVAS} tentativas: {ultima_excecao}"
        )

    # ------------------------------------------------------------------
    # Parsers de resposta
    # ------------------------------------------------------------------

    @staticmethod
    def _limpar_resposta(texto: str) -> str:
        texto = texto.strip().strip('"').strip("'")
        # Remove prefixos como "Categoria: " ou "Resposta: "
        for prefixo in ("categoria:", "resposta:", "resultado:"):
            if texto.lower().startswith(prefixo):
                texto = texto[len(prefixo):].strip()
        return texto or "SEM_RESPOSTA"

    @staticmethod
    def _parsear_semi(texto: str, cats_imp: list[str]) -> tuple[str, str]:
        """
        Extrai IMPUTADO e NOVO do formato de duas linhas.
        Retorna (imputado, novo).
        """
        imputado = ""
        novo = ""

        for linha in texto.splitlines():
            linha_l = linha.strip().lower()
            if linha_l.startswith("imputado:"):
                imputado = linha.split(":", 1)[1].strip()
            elif linha_l.startswith("novo:"):
                novo = linha.split(":", 1)[1].strip()

        # Fallback: se o modelo nao seguiu o formato, tenta interpretar
        if not imputado and not novo:
            # Verifica se bate com alguma categoria pre-definida
            cats_lower = {c.lower(): c for c in cats_imp}
            if texto.strip().lower() in cats_lower:
                imputado = cats_lower[texto.strip().lower()]
            else:
                imputado = "SEM_IMPUTACAO"
                novo = CodificadorIA._limpar_resposta(texto)

        # Normaliza SEM_IMPUTACAO
        if imputado.upper() in ("SEM_IMPUTACAO", "NENHUMA", "NONE", ""):
            imputado = "SEM_IMPUTACAO"

        return imputado, novo

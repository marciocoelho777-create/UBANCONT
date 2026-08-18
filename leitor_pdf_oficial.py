# -*- coding: utf-8 -*-
"""
Leitor dos demonstrativos contábeis OFICIAIS (PDFs publicados na pasta
"13 - DEMONSTRATIVOS CONTÁBEIS"), para o relatorio_integridade.py.

Substitui o antigo fluxo de entrada manual (dados_mensais.xlsx) por leitura
direta dos PDFs -- elimina o passo de digitação todo mês e permite rodar
sozinho na rotina noturna.

Não depende do Oracle: lê só os PDFs já publicados pela área de
contabilidade em "13 - DEMONSTRATIVOS CONTÁBEIS/{ano}/{subpasta}/
Lista{Tipo} {mes:02d}.pdf" -- os mesmos 6 demonstrativos oficiais
(BF, BP, DVP, BO, DFC, DMPL).
"""
from __future__ import annotations
import re
from pathlib import Path
from dataclasses import dataclass, field

import pdfplumber

RAIZ_PASTA13 = Path(__file__).parent / "13 - DEMONSTRATIVOS CONTÁBEIS"

# (subpasta, prefixo do arquivo) por tipo de demonstrativo, layout usado a
# partir de 2026 (com subpasta por tipo). Anos anteriores (ex.: 2025) usam
# os PDFs direto na raiz do ano, com sufixo "13" (encerramento) -- ver
# _localizar_pdf().
SUBPASTA_PREFIXO = {
    "BF":   ("01 - Balanço Financeiro",   "ListaBalancoFinanceiro"),
    "BP":   ("02 - Balanço Patrimonial",  "ListaBalancoPatrimonial"),
    "DVP":  ("03 - Variação Patrimonial", "ListaVariacaoPatrimonial"),
    "BO":   ("04 - Balanço Orçamentário", "ListaBalancoOrcamentario"),
    "DFC":  ("05 - Fluxo de Caixa",       "ListaFluxoCaixa"),
    "DMPL": ("06 - DMPL",                 "ListaDMPL"),
}


class PdfNaoEncontrado(Exception):
    pass


def _localizar_pdf(tipo: str, mes: int, ano: int) -> Path:
    subpasta, prefixo = SUBPASTA_PREFIXO[tipo]
    candidatos = [
        RAIZ_PASTA13 / str(ano) / subpasta / f"{prefixo} {mes:02d}.pdf",
        RAIZ_PASTA13 / str(ano) / f"{prefixo} {mes:02d}.pdf",
    ]
    if mes == 12:
        candidatos.append(RAIZ_PASTA13 / str(ano) / f"{prefixo} 13.pdf")
    for c in candidatos:
        if c.exists():
            return c
    raise PdfNaoEncontrado(
        f"Não encontrei o PDF de {tipo} para {mes:02d}/{ano} -- tentei: "
        + "; ".join(str(c) for c in candidatos))


# Páginas com 2 blocos lado a lado (ex.: BF pág.1 = Ingressos | Dispêndio)
# -- extract_text() padrão do pdfplumber intercala as linhas dos dois
# blocos quando eles não têm exatamente o mesmo número de linhas na mesma
# altura, embaralhando rótulo de um lado com valor do outro. Corta cada
# página no meio (x) e lê metade esquerda inteira, depois direita inteira.
PAGINAS_DUAS_COLUNAS = {
    "BF":  {0},  # página 1 (índice 0): Ingressos | Dispêndio
    "DVP": {0},  # página 1 (índice 0): VPA | VPD e Resultado Patrimonial
}


def _texto_completo(caminho: Path, tipo: str | None = None) -> str:
    duas_col = PAGINAS_DUAS_COLUNAS.get(tipo or "", set())
    partes = []
    with pdfplumber.open(caminho) as pdf:
        for idx, pg in enumerate(pdf.pages):
            if idx in duas_col:
                meio = pg.width / 2
                esq = pg.crop((0, 0, meio, pg.height)).extract_text() or ""
                dirt = pg.crop((meio, 0, pg.width, pg.height)).extract_text() or ""
                partes.append(esq)
                partes.append(dirt)
            else:
                partes.append(pg.extract_text() or "")
    return "\n".join(partes)


# NÃO inclui um "-" isolado como número válido: alguns rótulos vêm
# truncados no PDF com uma fórmula colada ("(III) = (I - II)" perde o
# "II)" na extração e sobra um "-" solto antes do valor real), e um "-"
# isolado nesse contexto é o sinal de subtração da fórmula, não uma
# célula vazia/zerada -- se fosse aceito, a busca por fallback pegaria
# esse "-" como o primeiro valor (0.0) em vez do valor real logo depois.
_NUM = r"(-?[\d\.]+,\d{2})"


def _num(s: str) -> float:
    s = s.strip()
    if s in ("-", "", "—"):
        return 0.0
    neg = s.startswith("-")
    s = s.lstrip("-").replace(".", "").replace(",", ".")
    v = float(s)
    return -v if neg else v


def valores(texto: str, label: str, n: int, inicio: int = 0):
    """Acha `label` em `texto` (a partir da posição `inicio`) e retorna os
    `n` números que vêm logo em seguida, como floats. Lança ValueError se
    não achar o label ou não achar `n` números depois dele."""
    idx = texto.find(label, inicio)
    if idx == -1:
        raise ValueError(f"Rótulo não encontrado: {label!r}")
    resto = texto[idx + len(label):]
    padrao = r"\s+".join([_NUM] * n)
    m = re.match(r"\s*" + padrao, resto)
    if not m:
        # fallback: os números podem não estar imediatamente colados
        # (label em uma "coluna" textual, números na próxima) -- procura
        # os N primeiros tokens numéricos nos próximos ~400 caracteres.
        janela = resto[:400]
        achados = re.findall(_NUM, janela)
        if len(achados) < n:
            raise ValueError(
                f"Não achei {n} valor(es) após o rótulo {label!r} "
                f"(achei {len(achados)}).")
        return [_num(v) for v in achados[:n]], idx
    return [_num(g) for g in m.groups()], idx


def pos(texto: str, marcador: str, a_partir_de: int = 0) -> int:
    idx = texto.find(marcador, a_partir_de)
    if idx == -1:
        raise ValueError(f"Marcador não encontrado: {marcador!r}")
    return idx


@dataclass
class Demonstrativos:
    mes: int
    ano: int
    caminhos: dict = field(default_factory=dict)
    textos: dict = field(default_factory=dict)

    def t(self, tipo: str) -> str:
        return self.textos[tipo]


def carregar(mes: int, ano: int) -> Demonstrativos:
    """Localiza e lê (texto bruto) os 6 PDFs oficiais do mês/ano pedido."""
    d = Demonstrativos(mes=mes, ano=ano)
    for tipo in SUBPASTA_PREFIXO:
        caminho = _localizar_pdf(tipo, mes, ano)
        d.caminhos[tipo] = caminho
        d.textos[tipo] = _texto_completo(caminho, tipo=tipo)
    return d


def ultimo_mes_disponivel(ano: int) -> int | None:
    """Maior mês (1-12) para o qual os 6 demonstrativos já estão
    publicados na pasta 13, ou None se nenhum mês estiver completo. Usado
    pela rotina noturna para não precisar saber de antemão até que mês o
    órgão já fechou/publicou."""
    for mes in range(12, 0, -1):
        try:
            for tipo in SUBPASTA_PREFIXO:
                _localizar_pdf(tipo, mes, ano)
        except PdfNaoEncontrado:
            continue
        return mes
    return None


def arquivos_existentes(ano: int):
    """Lista todos os arquivos PDF sob a pasta 13 para o ano dado (usado
    para detectar alteração/novo mês publicado antes de rodar o relatório
    -- ver PASSO de checagem no SKILL.md da rotina noturna)."""
    base = RAIZ_PASTA13 / str(ano)
    if not base.exists():
        return []
    return sorted(base.rglob("*.pdf"))


def fingerprint_pasta13(ano: int) -> str:
    """Hash (nome+tamanho+data de modificação de cada PDF) da pasta 13 do
    ano dado -- usado para saber se algo mudou (novo mês publicado, PDF
    substituído) desde a última rodada, sem reabrir/reler cada arquivo."""
    import hashlib
    h = hashlib.sha256()
    for p in arquivos_existentes(ano):
        st = p.stat()
        h.update(f"{p.relative_to(RAIZ_PASTA13)}|{st.st_size}|{int(st.st_mtime)}\n".encode("utf-8"))
    return h.hexdigest()


def _arquivo_estado(ano: int) -> Path:
    return Path(__file__).parent / f".pasta13_estado_{ano}.txt"


def pasta13_mudou(ano: int) -> bool:
    """True se a pasta 13 do ano dado mudou desde a última chamada a
    marcar_pasta13_processada() -- ou se nunca foi processada."""
    atual = fingerprint_pasta13(ano)
    arq = _arquivo_estado(ano)
    anterior = arq.read_text(encoding="utf-8").strip() if arq.exists() else None
    return atual != anterior


def marcar_pasta13_processada(ano: int) -> None:
    """Grava o fingerprint atual da pasta 13 -- chamar depois de gerar o
    relatório com sucesso, para a próxima rodada saber que já processou
    este estado da pasta."""
    _arquivo_estado(ano).write_text(fingerprint_pasta13(ano), encoding="utf-8")

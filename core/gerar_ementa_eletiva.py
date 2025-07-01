"""Módulo para geração de ementas eletivas em formato DOCX."""

from pathlib import Path
import base64
import logging
from typing import Optional, Union
import traceback
from docxtpl import DocxTemplate
from datetime import datetime

logger = logging.getLogger(__name__)

def gerar_ementa_eletiva(
    titulo: str,
    tema: str,
    professor1: str,
    justificativa: str,
    objetivo: str,
    professor2: str = "",
    habilidades: str = "",
    conteudo: str = "",
    metodologia: str = "",
    recursos: str = "",
    culminancia: str = "",
    referencia: str = "",
    ano_serie: str = "",
    return_base64: bool = True,
    base_path: Optional[Path] = None
) -> Union[str, dict]:
    """
    Gera documento de ementa eletiva no formato DOCX usando template base.
    
    Args:
        titulo: Título da eletiva (usado para nomear o arquivo)
        tema: Áreas/disciplinas envolvidas
        professor1: Nome do primeiro professor (obrigatório)
        justificativa: Justificativa da eletiva
        objetivo: Objetivos de aprendizagem
        professor2: Nome do segundo professor (opcional)
        habilidades: Habilidades desenvolvidas
        conteudo: Conteúdo programático
        metodologia: Métodos de ensino
        recursos: Recursos necessários
        culminancia: Atividade de culminância
        referencia: Referências bibliográficas
        ano_serie: Série/ano destinado (opcional)
        return_base64: Se True, retorna base64 do arquivo
        base_path: Caminho base do projeto (opcional)
    
    Returns:
        dict com status e arquivo em base64, ou str com caminho do arquivo
    Raises:
        FileNotFoundError: Quando o template não é encontrado
        Exception: Para outros erros durante a geração do documento
    """
    try:
        # Configura caminhos
        base_path = base_path or Path(__file__).parent.parent
        template_path = base_path / "complementos/templates/template_eletivas_2025.docx"
        
        # Valida template
        if not template_path.exists():
            raise FileNotFoundError(
                f"Template não encontrado em: {template_path}\n"
                "Verifique se o arquivo template_eletivas_2025.docx está na pasta templates."
            )
        
        # Gera nome do arquivo seguro
        nome_arquivo = (
            f"EMENTA_"
            f"{_sanitize_filename(titulo)}.docx"
        )
        output_path = base_path / "output" / nome_arquivo
        
        # Cria diretórios se necessário
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Prepara contexto para o template
        context = _prepare_context(
            titulo, tema, professor1, professor2, ano_serie,
            justificativa, objetivo, habilidades, conteudo,
            metodologia, recursos, culminancia, referencia
        )
        
        # Renderiza e salva documento
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)
        
        return _prepare_output(return_base64, output_path, nome_arquivo)
        
    except Exception as e:
        logger.error("Erro ao gerar ementa: %s\n%s", str(e), traceback.format_exc())
        raise

def _sanitize_filename(filename: str) -> str:
    """Remove caracteres inválidos para nomes de arquivo."""
    return (
        "".join(c if c.isalnum() or c in " _-()" else "_" for c in filename)
        .upper()
        .replace(" ", "_")
    )

def _prepare_context(
    titulo: str, tema: str, professor1: str, professor2: str, ano_serie: str,
    justificativa: str, objetivo: str, habilidades: str, conteudo: str,
    metodologia: str, recursos: str, culminancia: str, referencia: str
) -> dict:
    """Prepara o dicionário de contexto para o template."""
    return {
        "TITULO": titulo,
        "TEMA": tema,
        "PROFESSOR1": professor1,
        "PROFESSOR2": professor2,
        "ANO_SERIE": f" - {ano_serie}" if ano_serie else "",
        "JUSTIFICATIVA": justificativa,
        "OBJETIVO": objetivo,
        "HABILIDADES": habilidades,
        "CONTEUDO": conteudo,
        "METODOLOGIA": metodologia,
        "RECURSOS": recursos,
        "CULMINANCIA": culminancia,
        "REFERENCIA": referencia,
        "DATA_GERACAO": datetime.now().strftime("%d/%m/%Y %H:%M")
    }

def _prepare_output(return_base64: bool, output_path: Path, filename: str) -> Union[str, dict]:
    """Prepara o retorno conforme o formato solicitado."""
    if return_base64:
        with open(output_path, "rb") as f:
            return {
                "status": "success",
                "file_base64": base64.b64encode(f.read()).decode('utf-8'),
                "file_name": filename
            }
    return str(output_path)
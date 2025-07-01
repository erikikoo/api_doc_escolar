# api/main.py
from fastapi import FastAPI, HTTPException, Request, Query
from fastapi.responses import JSONResponse
import logging
import uvicorn
import traceback
import sys
from pathlib import Path
from typing import Optional
from core.gerar_ementa_eletiva import gerar_ementa_eletiva  # Você precisará criar esta função


# Configuração de logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuração dos caminhos
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.append(str(PROJECT_ROOT))

# Importações dos módulos core
from core.gerar_agenda import criar_agenda
from core.gerar_guias import gerar_guias

app = FastAPI()

@app.post("/webhook/n8n/gerar-agenda")
async def gerar_agenda_api(
    request: Request,
    mes: Optional[int] = Query(None),
    ano: Optional[int] = Query(None),
    professor: Optional[str] = Query(None),
    return_base64: Optional[bool] = Query(True)
):
    """Endpoint para geração de agendas"""
    try:
        # Tenta obter dados do JSON body primeiro
        try:
            data = await request.json()
            mes = data.get('mes', mes)
            ano = data.get('ano', ano)
            professor = data.get('professor', professor)
            return_base64 = data.get('return_base64', return_base64)
        except Exception as json_error:
            logger.debug(f"Falha ao ler JSON: {json_error}")

        if None in [mes, ano, professor]:
            raise HTTPException(
                status_code=400,
                detail="Parâmetros obrigatórios faltando: mes, ano e professor"
            )

        result = criar_agenda(
            mes=mes,
            ano=ano,
            professor=professor,
            return_base64=return_base64
        )
        
        if isinstance(result, dict):
            return result
        return {
            "status": "success",
            "file_url": str(result),
            "details": {
                "mes": mes,
                "ano": ano,
                "professor": professor
            }
        }
        
    except Exception as e:
        logger.error(f"Erro em gerar-agenda: {traceback.format_exc()}")
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(e),
                "type": type(e).__name__,
                "traceback": traceback.format_exc()
            }
        )

@app.post("/webhook/n8n/guias")
async def webhook_n8n_guias(
    request: Request,
    professor: Optional[str] = Query(None),
    disciplina: Optional[str] = Query(None),
    ano_serie: Optional[str] = Query(None),
    bimestre: Optional[str] = Query(None),
    ciclo: Optional[int] = Query(None),  # Alterado para int
    fontes: Optional[str] = Query(None)
):
    """Endpoint para integração com n8n"""
    try:
        # Tenta obter dados do JSON body primeiro
        try:
            data = await request.json()
            professor = data.get('professor', professor)
            disciplina = data.get('disciplina', disciplina)
            ano_serie = data.get('ano_serie', ano_serie)
            bimestre = data.get('bimestre', bimestre)
            ciclo = data.get('ciclo', ciclo)
            fontes = data.get('fontes', fontes)
        except Exception as json_error:
            logger.debug(f"Falha ao ler JSON, usando form data: {json_error}")
            data = await request.form()
            professor = professor or data.get('professor')
            disciplina = disciplina or data.get('disciplina')
            ano_serie = ano_serie or data.get('ano_serie')
            bimestre = bimestre or data.get('bimestre')
            ciclo = ciclo or data.get('ciclo')
            fontes = fontes or data.get('fontes')

        # Validação dos parâmetros
        missing = []
        if not professor: missing.append("professor")
        if not disciplina: missing.append("disciplina")
        if not ano_serie: missing.append("ano_serie")
        if not bimestre: missing.append("bimestre")
        if not ciclo: missing.append("ciclo")
        if not fontes: missing.append("fontes")
        
        if missing:
            raise HTTPException(
                status_code=400,
                detail=f"Parâmetros obrigatórios faltando: {', '.join(missing)}"
            )

        # Converter fontes de string para lista (se necessário)
        fontes_lista = fontes
        if isinstance(fontes, str):
            fontes_lista = [fonte.strip() for fonte in fontes.split(',') if fonte.strip()]

        result = gerar_guias(
            professor=professor,
            disciplina=disciplina,
            ano_serie=ano_serie,
            bimestre=bimestre,
            ciclo=int(ciclo),  # Garantir que é inteiro
            fontes=fontes_lista,  # Passar a lista de fontes
            base_path=PROJECT_ROOT,
            return_base64=True
        )
        
        return JSONResponse(result)
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro em webhook/n8n/guias: {traceback.format_exc()}")
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(e),
                "type": type(e).__name__,
                "traceback": traceback.format_exc(),
                "received_data": {
                    "professor": professor,
                    "disciplina": disciplina,
                    "ano_serie": ano_serie,
                    "bimestre": bimestre,
                    "ciclo": ciclo,
                    "fontes": fontes
                }
            }
        )

@app.post("/webhook/n8n/gerar-ementa-eletiva")
async def gerar_ementa_eletiva_api(
    request: Request,    
    titulo: Optional[str] = Query(None),
    tema: Optional[str] = Query(None),
    professor1: Optional[str] = Query(None),
    professor2: Optional[str] = Query(None),
    ano_serie: Optional[str] = Query(None),
    justificativa: Optional[str] = Query(None),
    objetivo: Optional[str] = Query(None),
    habilidades: Optional[str] = Query(None),
    conteudo: Optional[str] = Query(None),
    metodologia: Optional[str] = Query(None),
    recursos: Optional[str] = Query(None),
    culminancia: Optional[str] = Query(None),
    referencia: Optional[str] = Query(None),
    return_base64: Optional[bool] = Query(True)
):
    """Endpoint para geração de ementas eletivas"""
    try:
        # Tenta obter dados do JSON body primeiro
        try:
            data = await request.json()
            titulo = data.get('titulo', titulo)
            tema = data.get('tema', tema)
            professor1 = data.get('professor1', professor1) or data.get('professores', {}).get('professor1', professor1)
            professor2 = data.get('professor2', professor2) or data.get('professores', {}).get('professor2', professor2)
            ano_serie = data.get('ano_serie', ano_serie)
            justificativa = data.get('justificativa', justificativa)
            objetivo = data.get('objetivo', objetivo)
            habilidades = data.get('habilidades', habilidades)
            conteudo = data.get('conteudo', conteudo)
            metodologia = data.get('metodologia', metodologia)
            recursos = data.get('recursos', recursos)
            culminancia = data.get('culminancia', culminancia)
            referencia = data.get('referencia', referencia)
            return_base64 = data.get('return_base64', return_base64)
        except Exception as json_error:
            logger.debug(f"Falha ao ler JSON: {json_error}")

        # Validação dos parâmetros obrigatórios
        missing = []
        if not titulo: missing.append("titulo")
        if not tema: missing.append("tema")
        if not professor1: missing.append("professor1")
        if not justificativa: missing.append("justificativa")
        if not objetivo: missing.append("objetivo")
        
        if missing:
            raise HTTPException(
                status_code=400,
                detail=f"Parâmetros obrigatórios faltando: {', '.join(missing)}"
            )

        result = gerar_ementa_eletiva(
            titulo=titulo,
            tema=tema,
            professor1=professor1,
            professor2=professor2 or "",
            ano_serie=ano_serie or "",
            justificativa=justificativa,
            objetivo=objetivo,
            habilidades=habilidades or "",
            conteudo=conteudo or "",
            metodologia=metodologia or "",
            recursos=recursos or "",
            culminancia=culminancia or "",
            referencia=referencia or "",
            return_base64=return_base64,
            base_path=PROJECT_ROOT
        )
        
        if isinstance(result, dict):
            return result
        return {
            "status": "success",
            "file_url": str(result),
            "details": {
                "titulo": titulo,
                "tema": tema,
                "professores": {
                    "professor1": professor1,
                    "professor2": professor2
                },
                "ano_serie": ano_serie
            }
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro em gerar-ementa-eletiva: {traceback.format_exc()}")
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(e),
                "type": type(e).__name__,
                "traceback": traceback.format_exc()
            }
        )

@app.get("/healthcheck")
async def healthcheck():
    """Endpoint para verificação de saúde da API"""
    return {
        "status": "healthy",
        "version": "1.0.0",
        "services": {
            "gerar_agenda": True,
            "gerar_guias": True,
            "gerar_ementa_eletiva": True
        }
    }

if __name__ == "__main__":
    uvicorn.run(
        "api.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="debug",
        access_log=True
    )
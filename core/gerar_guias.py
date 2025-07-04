import pandas as pd
from docxtpl import DocxTemplate
import os
from pathlib import Path
import logging
from io import BytesIO
import base64
import unicodedata
import re

def normalizar_texto(texto):
    """Remove acentos e converte para minúsculas para comparação"""
    if pd.isna(texto) or texto is None:
        return ""
    try:
        texto = str(texto)
        texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
        return texto.lower().strip()
    except Exception as e:
        logging.warning(f"Erro ao normalizar texto: {texto} - {str(e)}")
        return ""

def extrair_numero_serie(texto):
    """Extrai o número da série/ano, removendo 'ano' e caracteres especiais"""
    if pd.isna(texto) or texto is None:
        return ""
    
    texto = str(texto).strip()
    # Remove 'ano' e variações
    texto = re.sub(r'ano\s*', '', texto, flags=re.IGNORECASE)
    # Extrai apenas números e símbolos de grau/ordinal
    match = re.search(r'(\d+)\s*[°ª]?\s*', texto)
    return match.group(1) if match else ""

def encontrar_coluna(df, padroes):
    """Encontra coluna que corresponde a qualquer um dos padrões"""
    for col in df.columns:
        try:
            col_normalizada = normalizar_texto(col)
            for padrao in padroes:
                if normalizar_texto(padrao) in col_normalizada:
                    return col
        except Exception as e:
            logging.warning(f"Erro ao processar coluna {col}: {str(e)}")
            continue
    return None

def filtrar_dataframe(df, ano_serie, bimestre):
    """Filtra o dataframe de forma robusta, com tratamento flexível para ano/série e bimestre"""
    try:
        # Cria cópias das colunas como strings
        df = df.copy()
        
        # Normaliza o ano/série de entrada
        numero_serie = extrair_numero_serie(ano_serie)
        if not numero_serie:
            raise ValueError(f"Formato inválido para Ano/Série: {ano_serie}")
        
        # Cria coluna temporária com números extraídos (mais flexível)
        df['_ano_temp'] = df['AnoSerie'].astype(str).apply(
            lambda x: extrair_numero_serie(x) or normalizar_texto(x)
        )
        
        # Cria máscara para o bimestre (aceita número ou texto completo)
        df['_bim_temp'] = df['Bimestre'].astype(str).apply(
            lambda x: extrair_numero_serie(x) or normalizar_texto(x)
        )
        
        # Prepara valor de busca para bimestre (aceita número ou texto)
        bimestre_busca = extrair_numero_serie(bimestre) or normalizar_texto(bimestre)
        
        # Aplica filtro para ano/série (comparando números ou texto normalizado)
        mask_ano = df['_ano_temp'].str.contains(numero_serie, na=False, regex=False)
        
        # Aplica filtro para bimestre (comparando números ou texto normalizado)
        mask_bim = df['_bim_temp'].str.contains(bimestre_busca, na=False, regex=False)
        
        # Aplica filtro combinado
        filtered_df = df[mask_ano & mask_bim].copy()
        
        # Remove colunas temporárias
        filtered_df.drop(['_ano_temp', '_bim_temp'], axis=1, inplace=True)
        
        # Filtra linhas com título não vazio
        result = filtered_df[filtered_df['Titulo'].astype(str).str.strip().ne('')]
        
        # Adiciona logs para depuração
        if result.empty:
            logging.warning(f"Nenhum dado encontrado com os filtros:")
            logging.warning(f"Ano/Série buscado: {numero_serie} | Valores únicos na planilha: {df['AnoSerie'].unique()}")
            logging.warning(f"Bimestre buscado: {bimestre_busca} | Valores únicos na planilha: {df['Bimestre'].unique()}")
        
        return result
        
    except Exception as e:
        raise ValueError(f"Erro ao filtrar dados: {str(e)}")

import json
import logging
from ast import literal_eval

import json
import logging
from ast import literal_eval

import json
import logging
from ast import literal_eval

import json
import logging
from ast import literal_eval

import json
import logging
import re

import json
import logging
import re
from ast import literal_eval

def formatar_fontes(fontes) -> str:
    """Formata as fontes para o template com tratamento completo para strings inválidas"""
    DEFAULT_RESPONSE = "• Materiais didáticos\n\n• Plataformas digitais\n\n• Orientação do professor"
    
    if not fontes:
        return DEFAULT_RESPONSE

    try:
        # ETAPA 1: Normalização radical da entrada
        if isinstance(fontes, str):
            # Remove múltiplas camadas de aspas e escapes
            fontes_clean = fontes.strip()
            while (fontes_clean.startswith('"') and fontes_clean.endswith('"')) or \
                  (fontes_clean.startswith("'") and fontes_clean.endswith("'")):
                fontes_clean = fontes_clean[1:-1]
            
            fontes_clean = fontes_clean.replace('\\"', '"').replace("\\'", "'")
            
            # Debug: Mostrar conteúdo após limpeza
            logging.debug(f"Conteúdo limpo: {fontes_clean}")
            
            # Caso especial: strings fragmentadas (como nos seus logs)
            if '\n' in fontes_clean or fontes_clean.count('{') > 1:
                # Reconstrói o JSON manualmente
                partes = re.findall(r'\{.*?\}', fontes_clean.replace('\n', ''))
                fontes_clean = f'[{",".join(partes)}]' if partes else '[]'
            
            # Tenta parsear como JSON
            try:
                dados_fontes = json.loads(fontes_clean)
            except json.JSONDecodeError:
                try:
                    # Fallback para avaliação segura
                    dados_fontes = literal_eval(fontes_clean)
                except:
                    logging.error("Falha ao decodificar fontes após múltiplas tentativas")
                    return DEFAULT_RESPONSE
        else:
            dados_fontes = fontes

        # ETAPA 2: Garantir que temos uma lista de dicionários
        if not isinstance(dados_fontes, list):
            dados_fontes = [dados_fontes] if isinstance(dados_fontes, dict) else []
        
        fontes_validas = []
        for item in dados_fontes:
            if isinstance(item, dict):
                fontes_validas.append(item)
            elif isinstance(item, str):
                try:
                    # Tenta parsear strings individuais
                    parsed = json.loads(item)
                    if isinstance(parsed, dict):
                        fontes_validas.append(parsed)
                except:
                    logging.warning(f"Ignorando item inválido: {item}")
                    continue
        
        # ETAPA 3: Formatação final
        itens_formatados = []
        for fonte in fontes_validas[:5]:  # Limita a 5 fontes
            try:
                nome = str(fonte.get('fonte_nome') or fonte.get('nome') or '').strip()
                if not nome:
                    continue
                
                item = f"• {nome}"
                
                descricao = str(fonte.get('descricao') or fonte.get('description') or '').strip()
                if descricao:
                    item += f"\n  Descrição: {descricao}"
                
                link = str(fonte.get('link') or fonte.get('url') or '').strip()
                if link:
                    item += f"\n  Link: {link}"
                
                itens_formatados.append(item)
            except Exception as e:
                logging.error(f"Erro ao formatar fonte: {str(e)}")
                continue

        return '\n\n'.join(itens_formatados) if itens_formatados else "• Nenhuma fonte disponível"

    except Exception as e:
        logging.error(f"ERRO CRÍTICO: {str(e)}", exc_info=True)
        return DEFAULT_RESPONSE

def gerar_guias(professor: str, disciplina: str, ano_serie: str, bimestre: str, ciclo: int, 
                base_path: Path = None, return_base64: bool = True, fontes = None) -> dict:
    """Gera guias de aprendizagem com tratamento robusto de dados"""
    try:
        # Configuração de caminhos
        base_path = base_path or Path(__file__).parent.parent
        template_path = base_path / "complementos/templates/template_guia_aprendizagem_2025.docx"        
        output_folder = base_path / "outputs/guias"

        if ciclo == 1:
            excel_path = base_path / "complementos/dados/1. Anos Iniciais - Escopo-sequência 2025.xlsx"
        elif ciclo == 2:
            excel_path = base_path / "complementos/dados/2. Anos Finais - Escopo-sequência 2025.xlsx"
        elif ciclo == 3:
            excel_path = base_path / "complementos/dados/3. Ensino Médio - Escopo-sequência 2025.xlsx"
       
        # Verificação de arquivos
        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado em: {template_path}")
        if not excel_path.exists():
            raise FileNotFoundError(f"Planilha não encontrada em: {excel_path}")
        
        os.makedirs(output_folder, exist_ok=True)

        # Carrega os dados
        try:
            df = pd.read_excel(excel_path, sheet_name=disciplina)
        except Exception as e:
            available_sheets = pd.ExcelFile(excel_path).sheet_names
            raise ValueError(f"Erro ao acessar aba '{disciplina}'. Abas disponíveis: {available_sheets}") from e

        # Mapeamento de colunas com fallback
        colunas_mapeadas = {
            'AnoSerie': encontrar_coluna(df, ['ANO/SÉRIE', 'ANO', 'SÉRIE', 'ANO SERIE']) or 'AnoSerie',
            'Bimestre': encontrar_coluna(df, ['BIMESTRE', 'BIM', 'PERÍODO']) or 'Bimestre',
            'Titulo': encontrar_coluna(df, ['TÍTULO DA AULA', 'TITULO', 'NOME DA AULA']) or 'Titulo',
            'Conteudo': encontrar_coluna(df, ['CONTEÚDO', 'CONTEUDO', 'MATÉRIA', 'ASSUNTO']) or 'Conteudo',
            'Objetivos': encontrar_coluna(df, ['OBJETIVOS', 'OBJETIVO', 'METAS']) or 'Objetivos'
        }

        # Renomeia colunas
        df = df.rename(columns={v: k for k, v in colunas_mapeadas.items() if v is not None})
        
        # Preenche valores NaN com string vazia
        for col in ['AnoSerie', 'Bimestre', 'Titulo', 'Conteudo', 'Objetivos']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str)
            else:
                raise ValueError(f"Coluna obrigatória '{col}' não encontrada")

        # Filtra os dados
        filtered_df = filtrar_dataframe(df, ano_serie, bimestre)
        
        if filtered_df.empty:
            raise ValueError(f"Nenhum dado encontrado para: Disciplina={disciplina}, Ano/Série (formatos aceitos: '6°', '6° ano', '6ª', etc)={ano_serie}, Bimestre={bimestre}")

        # Função para formatar seções
        def formatar_secao(dados):
            itens_vistos = set()
            itens_unicos = []
            
            for item in dados:
                item_limpo = str(item).strip()
                if item_limpo and item_limpo not in itens_vistos:
                    itens_vistos.add(item_limpo)
                    itens_unicos.append(item_limpo)
            
            return '\n\n'.join(f'• {item}' for item in itens_unicos) if itens_unicos else "Nenhum conteúdo disponível"

        # Prepara o contexto
        context = {
            'Professor': professor,
            'Disciplina': disciplina,
            'AnoSerie': ano_serie,
            'Bimestre': bimestre,
            'Titulo': formatar_secao(filtered_df['Titulo']),
            'Conteudo': formatar_secao(filtered_df['Conteudo']),
            'Objetivos': formatar_secao(filtered_df['Objetivos']),
            'Fontes': formatar_fontes(fontes)
        }

        # Gera o documento
        doc = DocxTemplate(template_path)
        doc.render(context)

        # Nome do arquivo
        nome_arquivo = (            
            f"Guia_{professor[:20].replace(' ', '_')}_"
            f"{str(ano_serie).replace(' ', '')}_"
            f"Bim{bimestre}_"
            f"{disciplina[:30].replace(' ', '_')}.docx"
        )
        output_path = output_folder / nome_arquivo

        # Retorna Base64 ou arquivo físico
        if return_base64:
            with BytesIO() as buffer:
                doc.save(buffer)
                buffer.seek(0)
                file_content = buffer.getvalue()
                
                if len(file_content) < 1024:
                    raise ValueError("Arquivo gerado é muito pequeno")
                
                return {
                    'status': 'success',
                    'data': {
                        'file_base64': base64.b64encode(file_content).decode('utf-8'),
                        'mime_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'file_name': nome_arquivo
                    }
                }
        else:
            doc.save(output_path)
            return {
                'status': 'success',
                'data': {
                    'file_path': str(output_path),
                    'file_name': nome_arquivo
                }
            }

    except Exception as e:
        logging.error(f"ERRO: {str(e)}", exc_info=True)
        return {
            'status': 'error',
            'message': str(e),
            'error_details': {
                'professor': professor,
                'disciplina': disciplina,
                'ano_serie': ano_serie,
                'bimestre': bimestre
            }
        }
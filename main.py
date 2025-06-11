import os
import re
import pandas as pd
from pathlib import Path
import PyPDF2
import fitz  # PyMuPDF
from typing import Dict, List, Optional
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PDFExtractor:
    def __init__(self):
        self.results = []

    def extract_text_from_pdf(self, pdf_path: str) -> List[str]:
        """
        Extrai texto de cada página do PDF usando PyMuPDF para melhor OCR
        """
        try:
            doc = fitz.open(pdf_path)
            pages_text = []

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text = page.get_text()
                pages_text.append(text)

            doc.close()
            return pages_text
        except Exception as e:
            logger.error(f"Erro ao extrair texto do PDF {pdf_path}: {str(e)}")
            return []

    def extract_recibo_info(self, text: str) -> Dict[str, str]:
        """
        Extrai informações do recibo (primeira página)
        """
        info = {
            'nome_pagante': '',
            'nome_aluno': '',
            'mes_competencia': '',
            'valor_recibo': '',
            'nome_escola': '',
            'cnpj_escola': ''
        }

        try:
            # Extrair nome da escola (primeira linha ou linha com CNPJ)
            escola_match = re.search(r'^(.+?)(?:\s*CNPJ|\s*\n)', text, re.MULTILINE)
            if escola_match:
                info['nome_escola'] = escola_match.group(1).strip()

            # Extrair CNPJ da escola
            cnpj_match = re.search(r'CNPJ[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', text)
            if cnpj_match:
                info['cnpj_escola'] = cnpj_match.group(1)

            # Extrair nome do pagante
            pagante_patterns = [
                r'recebeu do\(?a?\)?\s*Sr\(?a?\)?\.\s*([^,]+)',
                r'Sr\(?a?\)?\.\s*([^,]+?)(?:,|\s*registrado)',
                r'do\(?a?\)?\s*Sr\(?a?\)?\.\s*([^,\n]+)'
            ]

            for pattern in pagante_patterns:
                pagante_match = re.search(pattern, text, re.IGNORECASE)
                if pagante_match:
                    info['nome_pagante'] = pagante_match.group(1).strip()
                    break

            # Extrair valor
            valor_patterns = [
                r'valor de\s*R\$\s*([\d.,]+)',
                r'R\$\s*([\d.,]+)',
                r'(\d+\.?\d*,\d{2})\s*\([^)]+\)'
            ]

            for pattern in valor_patterns:
                valor_match = re.search(pattern, text)
                if valor_match:
                    info['valor_recibo'] = valor_match.group(1).strip()
                    break

            # Extrair mês/competência
            mes_patterns = [
                r'referente\s+(?:a|à)\s+[^d]*de\s+(\w+\s+\d{4})',
                r'Mensalidade\s+de\s+(\w+\s+\d{4})',
                r'competência\s+(\w+\s+\d{4})',
                r'mês\s+de\s+(\w+\s+\d{4})'
            ]

            for pattern in mes_patterns:
                mes_match = re.search(pattern, text, re.IGNORECASE)
                if mes_match:
                    info['mes_competencia'] = mes_match.group(1).strip()
                    break

            # Extrair nome do aluno
            aluno_patterns = [
                r'Nome do aluno\s*\(?a?\)?:\s*([^,\n]+)',
                r'aluno\s*\(?a?\)?:\s*([^,\n]+)',
                r'Aluno\s*\(?a?\)?:\s*([^,\n]+)'
            ]

            for pattern in aluno_patterns:
                aluno_match = re.search(pattern, text, re.IGNORECASE)
                if aluno_match:
                    info['nome_aluno'] = aluno_match.group(1).strip()
                    break

        except Exception as e:
            logger.error(f"Erro ao extrair informações do recibo: {str(e)}")

        return info

    def extract_cartao_cnpj_info(self, text: str) -> Dict[str, str]:
        """
        Extrai informações do cartão CNPJ (segunda página)
        """
        info = {
            'numero_cnpj': '',
            'nome_escola_cnpj': '',
            'atividade_economica': '',
            'data_emissao': ''
        }

        try:
            # Extrair número CNPJ
            cnpj_patterns = [
                r'CNPJ[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
                r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
                r'CNPJ[:\s]*(\d{14})'
            ]

            for pattern in cnpj_patterns:
                cnpj_match = re.search(pattern, text)
                if cnpj_match:
                    info['numero_cnpj'] = cnpj_match.group(1)
                    break

            # Extrair nome da escola/razão social
            nome_patterns = [
                r'Razão Social[:\s]*([^\n]+)',
                r'Nome Empresarial[:\s]*([^\n]+)',
                r'RAZÃO SOCIAL[:\s]*([^\n]+)'
            ]

            for pattern in nome_patterns:
                nome_match = re.search(pattern, text, re.IGNORECASE)
                if nome_match:
                    info['nome_escola_cnpj'] = nome_match.group(1).strip()
                    break

            # Extrair atividade econômica
            atividade_patterns = [
                r'Atividade Econômica[:\s]*([^\n]+)',
                r'CNAE[:\s]*([^\n]+)',
                r'Atividade Principal[:\s]*([^\n]+)'
            ]

            for pattern in atividade_patterns:
                atividade_match = re.search(pattern, text, re.IGNORECASE)
                if atividade_match:
                    info['atividade_economica'] = atividade_match.group(1).strip()
                    break

            # Extrair data de emissão
            data_patterns = [
                r'Data de Emissão[:\s]*(\d{2}/\d{2}/\d{4})',
                r'Emitido em[:\s]*(\d{2}/\d{2}/\d{4})',
                r'Data[:\s]*(\d{2}/\d{2}/\d{4})'
            ]

            for pattern in data_patterns:
                data_match = re.search(pattern, text, re.IGNORECASE)
                if data_match:
                    info['data_emissao'] = data_match.group(1)
                    break

        except Exception as e:
            logger.error(f"Erro ao extrair informações do cartão CNPJ: {str(e)}")

        return info

    def process_single_pdf(self, pdf_path: str) -> Dict[str, str]:
        """
        Processa um único arquivo PDF
        """
        logger.info(f"Processando: {pdf_path}")

        # Extrair texto de todas as páginas
        pages_text = self.extract_text_from_pdf(pdf_path)

        if len(pages_text) < 2:
            logger.warning(f"PDF {pdf_path} não tem 2 páginas. Páginas encontradas: {len(pages_text)}")
            return {}

        # Extrair informações da primeira página (recibo)
        recibo_info = self.extract_recibo_info(pages_text[0])

        # Extrair informações da segunda página (cartão CNPJ)
        cnpj_info = self.extract_cartao_cnpj_info(pages_text[1])

        # Combinar informações
        combined_info = {
            'arquivo': os.path.basename(pdf_path),
            'caminho_completo': pdf_path,
            **recibo_info,
            **cnpj_info
        }

        return combined_info

    def process_directory(self, directory_path: str) -> List[Dict[str, str]]:
        """
        Processa todos os PDFs em um diretório
        """
        directory = Path(directory_path)
        pdf_files = list(directory.glob("*.pdf"))

        logger.info(f"Encontrados {len(pdf_files)} arquivos PDF")

        results = []
        failed_files = []

        for pdf_file in pdf_files:
            try:
                result = self.process_single_pdf(str(pdf_file))
                if result:
                    results.append(result)
                else:
                    failed_files.append(str(pdf_file))
            except Exception as e:
                logger.error(f"Erro ao processar {pdf_file}: {str(e)}")
                failed_files.append(str(pdf_file))

        logger.info(f"Processamento concluído. Sucessos: {len(results)}, Falhas: {len(failed_files)}")

        if failed_files:
            logger.warning(f"Arquivos que falharam: {failed_files}")

        return results

    def save_to_excel(self, data: List[Dict[str, str]], output_path: str):
        """
        Salva os dados extraídos em um arquivo Excel
        """
        if not data:
            logger.warning("Nenhum dado para salvar")
            return

        df = pd.DataFrame(data)

        # Reordenar colunas para melhor visualização
        column_order = [
            'arquivo', 'nome_pagante', 'nome_aluno', 'mes_competencia', 'valor_recibo',
            'nome_escola', 'cnpj_escola', 'numero_cnpj', 'nome_escola_cnpj',
            'atividade_economica', 'data_emissao', 'caminho_completo'
        ]

        # Reorganizar colunas (manter apenas as que existem)
        existing_columns = [col for col in column_order if col in df.columns]
        df = df[existing_columns]

        # Salvar em Excel
        df.to_excel(output_path, index=False)
        logger.info(f"Dados salvos em: {output_path}")

        # Mostrar estatísticas
        logger.info(f"Total de registros: {len(df)}")
        logger.info(f"Colunas: {list(df.columns)}")


def main():
    """
    Função principal para executar o extrator
    """
    # Configurar caminhos
    input_directory = input("Digite o caminho da pasta com os PDFs: ").strip()
    output_file = input("Digite o nome do arquivo Excel de saída (ex: resultados.xlsx): ").strip()

    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Verificar se o diretório existe
    if not os.path.exists(input_directory):
        print(f"Erro: Diretório '{input_directory}' não encontrado!")
        return

    # Criar instância do extrator
    extractor = PDFExtractor()

    # Processar todos os PDFs
    results = extractor.process_directory(input_directory)

    # Salvar resultados
    if results:
        extractor.save_to_excel(results, output_file)
        print(f"\nProcessamento concluído!")
        print(f"Resultados salvos em: {output_file}")
        print(f"Total de arquivos processados: {len(results)}")
    else:
        print("Nenhum arquivo foi processado com sucesso.")


if __name__ == "__main__":
    main()
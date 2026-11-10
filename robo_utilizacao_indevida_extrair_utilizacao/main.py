import logging
import os
from datetime import datetime
from pathlib import Path
from time import sleep

import pandas as pd
from dotenv import load_dotenv
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from sei_automacao.auth import efetuar_login
from sei_automacao.core.iframes import trocar_iframe
from sei_automacao.documento import inserir_bloco_ass
from sei_automacao.processo import incluir_doc_externo, incluir_doc_sei_memo
from sei_automacao.ubiquo.busca import acessar_processo
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

from robo_utilizacao_indevida_extrair_utilizacao.txts_memorando import (
    txt_sem_uso,
    txt_uso_dep,
    txt_uso_tit,
    txt_uso_tit_dep,
)

load_dotenv()

SEI_LOGIN = os.getenv('SEI_LOGIN')
SEI_SENHA = os.getenv('SEI_SENHA')
SEI_ORGAO = os.getenv('SEI_ORGAO')

AGORA = datetime.now().strftime('%Y_%m_%d_%H%M%S')
HOJE = datetime.today().strftime('%d_%m_%Y')
DATA_PESQUISA = '04/09/2025'


DIR_BASE = Path(__file__).resolve().parent.parent
DIR_DADOS = DIR_BASE / 'dados'
DIR_LOGS = DIR_BASE / 'logs'

DIR_INPUTS = DIR_DADOS / 'inputs'
DIR_UTILIZACOES_EXCEL = DIR_DADOS / 'utilizacoes' / 'excel'
DIR_UTILIZACOES_PDF = DIR_DADOS / 'utilizacoes' / 'pdf'

ARQ_CONSOLIDACAO = DIR_INPUTS / 'consolidacao.xlsx'
ARQ_MANTIS_TIT = DIR_INPUTS / 'MANTIS.8769 - Titulares_revisado.xlsx'
ARQ_MANTIS_DEP = DIR_INPUTS / 'MANTIS.8769 - Dependentes_revisado.xlsx'
ARQ_LOG = DIR_LOGS / f'{AGORA}_execucao.log'

COL_WITDHS_DEP_DICT = {
    'PRESTADOR': 180,
    'MATRICULA DEPENDENTE': 90,
    'NOME DEPENDENTE': 180,
    'DATA DO ATENDIMENTO': 90,
    'VALOR PAGO': 80,
}

COL_WITDHS_TIT_DICT = {
    'PRESTADOR': 300,
    'DATA DO ATENDIMENTO': 90,
    'PROCEDIMENTO': 300,
    'VALOR PAGO': 80,
}


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(ARQ_LOG, encoding='utf-8'),
        logging.StreamHandler(),
    ],
)

logger = logging.getLogger(__name__)


def pesquisa_uso_tit(tit_matricula, lista_datas):
    try:
        tit_path_planilha = (
            DIR_UTILIZACOES_EXCEL
            / f'Utilizacao-Titular-{tit_matricula}-{HOJE}.xlsx'
        )

        df = pd.read_excel(ARQ_MANTIS_TIT, 'Exportar Planilha')
        df = df.rename(
            columns={
                'NOME PRESTADOR': 'PRESTADOR',
                'VALOR A PAGAR': 'VALOR PAGO',
            }
        )

        df['DATA DO ATENDIMENTO'] = pd.to_datetime(
            df['DATA DO ATENDIMENTO'], errors='coerce', dayfirst=True
        )

        df_uso_tit = df[df['MATRICULA TITULAR'] == tit_matricula]

        datas_dt = [pd.to_datetime(d, dayfirst=True) for d in lista_datas]

        mascara = False
        for i in range(0, len(datas_dt), 2):
            data_ini = datas_dt[i]
            data_fim = datas_dt[i + 1]
            mascara |= (df_uso_tit['DATA DO ATENDIMENTO'] >= data_ini) & (
                df_uso_tit['DATA DO ATENDIMENTO'] <= data_fim
            )

        df_uso_tit = df_uso_tit[mascara]
        df_uso_tit['DATA DO ATENDIMENTO'] = df_uso_tit[
            'DATA DO ATENDIMENTO'
        ].dt.strftime('%d/%m/%Y')

        tit_possui_uso = not df_uso_tit.empty
        df_uso_tit.to_excel(tit_path_planilha, index=False)

        tit_soma_uso = df_uso_tit['VALOR PAGO'].sum()

        return tit_possui_uso, tit_path_planilha, tit_soma_uso, df_uso_tit

    except Exception:
        logger.exception('Erro ao pesquisar utilização do titular')
        raise


def pesquisa_uso_dep(tit_matricula, lista_datas):
    try:
        df = pd.read_excel(ARQ_MANTIS_DEP, 'Exportar Planilha')
        df = df.rename(
            columns={
                'NOME PRESTADOR': 'PRESTADOR',
                'VALOR A PAGAR': 'VALOR PAGO',
            }
        )

        df['DATA DO ATENDIMENTO'] = pd.to_datetime(
            df['DATA DO ATENDIMENTO'], errors='coerce', dayfirst=True
        )

        df_uso_dep = df[df['MATRICULA TITULAR'] == tit_matricula]

        datas_dt = [pd.to_datetime(d, dayfirst=True) for d in lista_datas]

        mascara = False
        for i in range(0, len(datas_dt), 2):
            data_ini = datas_dt[i]
            data_fim = datas_dt[i + 1]
            mascara |= (df_uso_dep['DATA DO ATENDIMENTO'] >= data_ini) & (
                df_uso_dep['DATA DO ATENDIMENTO'] <= data_fim
            )

        df_uso_dep = df_uso_dep[mascara]
        df_uso_dep['DATA DO ATENDIMENTO'] = df_uso_dep[
            'DATA DO ATENDIMENTO'
        ].dt.strftime('%d/%m/%Y')

        dep_possui_uso = not df_uso_dep.empty
        dep_path_planilha = (
            DIR_UTILIZACOES_EXCEL
            / f'Utilizacao-Dependentes-{tit_matricula}-{HOJE}.xlsx'
        )
        df_uso_dep.to_excel(dep_path_planilha, index=False)

        dep_soma_uso = df_uso_dep['VALOR PAGO'].sum()

        return dep_possui_uso, dep_path_planilha, dep_soma_uso, df_uso_dep

    except Exception:
        logger.exception('Erro ao pesquisar utilização dos dependentes')
        raise


def checa_label_utilizacao_inserida(driver: webdriver.Chrome) -> bool:
    try:
        driver.find_element(
            By.XPATH, "//img[contains(@title, 'UTILIZAÇÃO INSERIDA')]"
        )
        return True
    except Exception:
        return False


def checa_label_feito(driver: webdriver.Chrome) -> bool:
    try:
        driver.find_element(
            By.XPATH, "//img[contains(@title, 'Feito - DESATIVADO')]"
        )
        return True
    except Exception:
        return False


def _formatar_valor_monetario(valor: float) -> str:
    formatado = f'{valor:,.2f}'
    return formatado.replace(',', 'X').replace('.', ',').replace('X', '.')


def _build_table_row(row, valor_col_index, matricula_dep_index, normal):
    linha = []
    for i, cell in enumerate(row):
        if i == valor_col_index and pd.notna(cell):
            linha.append(_formatar_valor_monetario(float(cell)))
        elif i == matricula_dep_index and pd.notna(cell):
            linha.append(str(int(cell)))
        else:
            linha.append(Paragraph(str(cell), normal))
    return linha


def transformar_df_em_pdf(  # noqa: PLR0914
    tit_matricula, tit_nome, tipo, df: pd.DataFrame
) -> Path:
    styles = getSampleStyleSheet()
    normal = styles['BodyText']
    header_style = styles['Heading4']

    table_header = ParagraphStyle(
        'table_header',
        parent=styles['BodyText'],
        textColor=colors.white,
        fontName='Helvetica-Bold',
        alignment=1,
    )

    caminho_pdf = (
        DIR_UTILIZACOES_PDF / f'Utilizacao-{tipo}-{tit_matricula}.pdf'
    )

    doc = SimpleDocTemplate(
        str(caminho_pdf),
        pagesize=landscape(A4),
        leftMargin=30,
        rightMargin=30,
        topMargin=40,
        bottomMargin=30,
    )

    elements = []

    logo = Image('img/logo.jpg', width=120, height=60)
    info = Paragraph(
        f'<b>Relatório de Utilização Por {tipo}</b><br/>'
        f'Beneficiário Titular: {tit_nome} - {tit_matricula}',
        header_style,
    )
    elements.append(Table([[logo, info]], colWidths=[150, 650]))
    elements.append(Spacer(1, 20))

    valor_col_index = (
        df.columns.get_loc('VALOR PAGO')
        if 'VALOR PAGO' in df.columns
        else None
    )
    matricula_dep_index = (
        df.columns.get_loc('MATRICULA DEPENDENTE')
        if 'MATRICULA DEPENDENTE' in df.columns
        else None
    )

    data = [[Paragraph(col, table_header) for col in df.columns]]
    for _, row in df.iterrows():
        data.append(
            _build_table_row(row, valor_col_index, matricula_dep_index, normal)
        )

    if tipo == 'Titular':
        col_widths = [COL_WITDHS_TIT_DICT.get(col, 100) for col in df.columns]
    elif tipo == 'Dependente(s)':
        col_widths = [COL_WITDHS_DEP_DICT.get(col, 100) for col in df.columns]
    else:
        col_widths = []

    table = Table(data, colWidths=col_widths, repeatRows=1)

    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightblue]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]

    if valor_col_index is not None:
        style_commands.append((
            'ALIGN',
            (valor_col_index, 1),
            (valor_col_index, -1),
            'RIGHT',
        ))
        style_commands.append((
            'ALIGN',
            (valor_col_index, 0),
            (valor_col_index, 0),
            'RIGHT',
        ))

    table.setStyle(TableStyle(style_commands))
    elements.append(table)
    doc.build(elements)

    return caminho_pdf


def _incluir_docs_titular(
    driver, num_processo, tit_matricula, tit_nome, lista_datas
) -> tuple[bool, float]:
    logger.info(f'PROCESSO {num_processo}: Pesquisando utilização do TITULAR')
    tit_possui_uso, _, tit_soma_uso, df_uso_tit = pesquisa_uso_tit(
        tit_matricula, lista_datas
    )
    logger.info(
        f'PROCESSO {num_processo}: TITULAR — '
        f'possui uso: {tit_possui_uso}, total: R$ {tit_soma_uso:.2f}'
    )

    if not tit_possui_uso:
        return False, 0.0

    df_uso_tit = df_uso_tit.drop(
        columns=[
            'REFERENCIA',
            'COD_PRESTADOR',
            'NUM_REMESSA',
            'MATRICULA TITULAR',
            'NOME TITULAR',
            'CODIGO',
        ]
    )
    caminho_pdf_tit = transformar_df_em_pdf(
        tit_matricula, tit_nome, 'Titular', df_uso_tit
    )
    logger.info(
        f'PROCESSO {num_processo}: '
        'Incluindo documento externo — utilização do titular'
    )
    incluir_doc_externo(
        driver,
        tipo_doc='PDF',
        num=f'Utilização Tit. - {tit_matricula} - {AGORA}',
        formato='Nato-digital',
        data=HOJE,
        path_anexo=caminho_pdf_tit,
        nivel_acesso='Restrito',
        hipotese_legal='Informação Pessoal (Art. 31 da Lei nº 12.527/2011)',
    )
    return True, tit_soma_uso


def _incluir_docs_dependentes(
    driver, num_processo, tit_matricula, tit_nome, lista_datas
) -> tuple[bool, float]:
    logger.info(
        f'PROCESSO {num_processo}: Pesquisando utilização dos DEPENDENTES'
    )
    dep_possui_uso, _, dep_soma_uso, df_uso_dep = pesquisa_uso_dep(
        tit_matricula, lista_datas
    )
    logger.info(
        f'PROCESSO {num_processo}: DEPENDENTES — '
        f'possui uso: {dep_possui_uso}, total: R$ {dep_soma_uso:.2f}'
    )

    if not dep_possui_uso:
        return False, 0.0

    df_uso_dep = df_uso_dep.drop(
        columns=[
            'REFERENCIA',
            'COD_PRESTADOR',
            'NUM_REMESSA',
            'MATRICULA TITULAR',
            'NOME TITULAR',
            'CODIGO',
            'PROCEDIMENTO',
        ]
    )
    df_uso_dep = df_uso_dep[
        [
            'MATRICULA DEPENDENTE',
            'NOME DEPENDENTE',
            'PRESTADOR',
            'DATA DO ATENDIMENTO',
            'VALOR PAGO',
        ]
    ]
    caminho_pdf_dep = transformar_df_em_pdf(
        tit_matricula, tit_nome, 'Dependente(s)', df_uso_dep
    )
    logger.info(
        f'PROCESSO {num_processo}: '
        'Incluindo documento externo — utilização dos dependentes'
    )
    incluir_doc_externo(
        driver,
        tipo_doc='PDF',
        num=f'Utilização Dep. - {tit_matricula} - {AGORA}',
        formato='Nato-digital',
        data=HOJE,
        path_anexo=caminho_pdf_dep,
        nivel_acesso='Restrito',
        hipotese_legal='Informação Pessoal (Art. 31 da Lei nº 12.527/2011)',
    )
    sleep(5)
    return True, dep_soma_uso


def _montar_texto_memorando(
    tit_nome,
    tit_matricula,
    periodos_str,
    soma_uso_tot,
    uso_flags: tuple[bool, bool],
) -> str:
    tit_possui_uso, dep_possui_uso = uso_flags
    if not tit_possui_uso and not dep_possui_uso:
        return txt_sem_uso(
            tit_nome, tit_matricula, periodos_str, DATA_PESQUISA
        )
    if tit_possui_uso and not dep_possui_uso:
        return txt_uso_tit(
            tit_nome, tit_matricula, periodos_str, soma_uso_tot, DATA_PESQUISA
        )
    if tit_possui_uso and dep_possui_uso:
        return txt_uso_tit_dep(
            tit_nome, tit_matricula, periodos_str, soma_uso_tot, DATA_PESQUISA
        )
    return txt_uso_dep(
        tit_nome, tit_matricula, periodos_str, soma_uso_tot, DATA_PESQUISA
    )


def processar_processo(driver, processo):
    num_processo = processo['PROCESSO SEI']
    tit_nome: str = str(processo['NOME TITULAR']).upper()
    tit_matricula = processo['MATRICULA TITULAR']
    tem_dependente = processo['DEPENDENTE?'] == 'S'

    lista_datas = [
        datetime.strptime(d, '%d/%m/%Y').strftime('%d/%m/%Y')
        for d in processo['DATAS'].split(';')
    ]
    periodos = [
        f'{lista_datas[i]} a {lista_datas[i + 1]}'
        for i in range(0, len(lista_datas), 2)
    ]
    periodos_str = ', '.join(periodos)

    logger.info(f'PROCESSO {num_processo}: Iniciando processamento')

    acessar_processo(driver, num_processo)
    trocar_iframe(driver, 'ifrArvore')

    if checa_label_utilizacao_inserida(driver):
        logger.info(
            f'PROCESSO {num_processo}: '
            "Marcador 'UTILIZAÇÃO INSERIDA' encontrado — pulando"
        )
        return

    if checa_label_feito(driver):
        logger.info(
            f'PROCESSO {num_processo}: '
            "Marcador 'Feito - DESATIVADO' encontrado — pulando"
        )
        return

    tit_possui_uso, tit_soma_uso = _incluir_docs_titular(
        driver, num_processo, tit_matricula, tit_nome, lista_datas
    )
    soma_uso_tot = tit_soma_uso

    dep_possui_uso = False
    if tem_dependente:
        dep_possui_uso, dep_soma_uso = _incluir_docs_dependentes(
            driver, num_processo, tit_matricula, tit_nome, lista_datas
        )
        soma_uso_tot += dep_soma_uso

    texto_principal = _montar_texto_memorando(
        tit_nome,
        tit_matricula,
        periodos_str,
        soma_uso_tot,
        uso_flags=(tit_possui_uso, dep_possui_uso),
    )

    logger.info(f'PROCESSO {num_processo}: Incluindo memorando')
    incluir_doc_sei_memo(
        driver,
        vocativo='Prezados',
        destinatario_nome=('Assessoria de Políticas e Informações da Saúde'),
        assunto=' Retorno acerca dos serviços assistenciais utilizados',
        texto_principal=texto_principal,
        nivel_acesso='Público',
    )

    logger.info(f'PROCESSO {num_processo}: Inserindo bloco de assinatura')
    inserir_bloco_ass(driver, '7222079')

    # adicionar_marcador(driver, "UTILIZAÇÃO INSERIDA")

    logger.info(
        f'PROCESSO {num_processo}: '
        f'Processamento concluído — total: R$ {soma_uso_tot:.2f}'
    )


def main():
    try:
        logger.info('----- INICIO DO SCRIPT -----')
        os.environ['NO_PROXY'] = 'localhost,127.0.0.1'

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()
        driver.get('https://www.sei.mg.gov.br')

        efetuar_login(driver, SEI_LOGIN, SEI_SENHA, SEI_ORGAO)
        sleep(2)
        # selecionar_unidade(driver, "Gerência de Auditoria de Contas da Saúde - Utilização")  # noqa: E501

        df_processos = pd.read_excel(ARQ_CONSOLIDACAO, 'Sheet1')

        for _, processo in df_processos.iterrows():
            try:
                processar_processo(driver, processo)
            except Exception as e:
                num = processo['PROCESSO SEI']
                logger.error(
                    f'PROCESSO {num}: Erro no processamento: {e}',
                    exc_info=True,
                )
                raise

        logger.info('----- FIM DO SCRIPT -----')

    except Exception:
        logger.error('Erro na execução do main', exc_info=True)
        raise

    finally:
        driver.quit()


if __name__ == '__main__':
    main()

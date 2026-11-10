from sei_automacao.utils.numeros import num_extenso_formatado

_CSS = 'Texto_Justificado_Recuo_Primeira_Linha'


def par_1():
    return (  # noqa: E501
        'Considerando o período e o volume de atendimentos vinculados a'
        ' beneficiários que não possuíam vínculo ativo ou contribuição'
        ' vigente, esclarecemos que foi solicitada à Assessoria de'
        ' Governança Eletrônica  ASGE, por meio do Mantis nº 8769, a'
        ' consolidação dos serviços de assistência à saúde registrados para'
        ' beneficiários nessa situação e devidamente faturados pelos'
        ' prestadores da Rede Assistencial do Instituto.'
    )


def par_2(data_pesquisa):
    return (
        f'O levantamento contempla a relação de procedimentos faturados e'
        f' exportados para o Sistema de Assistência Descentralizada à Saúde'
        f' - SADS até {data_pesquisa} e que, após validação pelo Instituto,'
        f' encontram-se em situação de auditoria concluída, liberação para'
        f' pagamento, processamento ou pagamento efetivado.'
    )


def txt_sem_uso(tit_nome, tit_matricula, periodos_str, data_pesquisa):
    return (
        f'<p class="{_CSS}">{par_1()}</p>'
        f'<p class="{_CSS}">{par_2(data_pesquisa)}</p>'
        f'<p class="{_CSS}">Em atenção à demanda encaminhada no presente'
        f' processo, informamos que não constam atendimentos prestados ao(à)'
        f' beneficiário(a) <strong>{tit_nome}</strong>, matrícula'
        f' <strong>{tit_matricula}</strong>,'
        f' no período de {periodos_str}.</p>'
    )


def txt_uso_tit(
    tit_nome, tit_matricula, periodos_str, soma_uso_tot, data_pesquisa
):
    valor_fmt, int_ext, dec_ext = num_extenso_formatado(soma_uso_tot)
    return (
        f'<p class="{_CSS}">{par_1()}</p>'
        f'<p class="{_CSS}">{par_2(data_pesquisa)}</p>'
        f'<p class="{_CSS}">Em atenção à demanda encaminhada no presente'
        f' processo, informamos que constam atendimentos prestados ao(à)'
        f' beneficiário(a) <strong>{tit_nome}</strong>, matrícula'
        f' <strong>{tit_matricula}</strong>, no período de {periodos_str},'
        f' totalizando o valor de <strong>R$ {valor_fmt}</strong>'
        f' ({int_ext} reais e {dec_ext} centavos).</p>'
    )


def txt_uso_tit_dep(
    tit_nome, tit_matricula, periodos_str, soma_uso_tot, data_pesquisa
):
    valor_fmt, int_ext, dec_ext = num_extenso_formatado(soma_uso_tot)
    return (
        f'<p class="{_CSS}">{par_1()}</p>'
        f'<p class="{_CSS}">{par_2(data_pesquisa)}</p>'
        f'<p class="{_CSS}">Em atenção à demanda encaminhada no presente'
        f' processo, informamos que constam atendimentos prestados ao(à)'
        f' beneficiário(a) <strong>{tit_nome}</strong>, matrícula'
        f' <strong>{tit_matricula}</strong>, e a seus dependentes, no'
        f' período de {periodos_str}, totalizando o valor de'
        f' <strong>R$ {valor_fmt}</strong>'
        f' ({int_ext} reais e {dec_ext} centavos).</p>'
    )


def txt_uso_dep(
    tit_nome, tit_matricula, periodos_str, soma_uso_tot, data_pesquisa
):
    valor_fmt, int_ext, dec_ext = num_extenso_formatado(soma_uso_tot)
    return (
        f'<p class="{_CSS}">{par_1()}</p>'
        f'<p class="{_CSS}">{par_2(data_pesquisa)}</p>'
        f'<p class="{_CSS}">Em atenção à demanda encaminhada no presente'
        f' processo, informamos que constam atendimentos prestados aos'
        f' dependentes do beneficiário(a) <strong>{tit_nome}</strong>,'
        f' matrícula <strong>{tit_matricula}</strong>, no período de'
        f' {periodos_str}, totalizando o valor de'
        f' <strong>R$ {valor_fmt}</strong>'
        f' ({int_ext} reais e {dec_ext} centavos).</p>'
    )

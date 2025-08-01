# Library import

from openpyxl import load_workbook
from solutions.adat.general_defs import (

    files_import, 
    id_normalize_apply, 
    fy_dates, 
    fy_filter,
    actives_false_positive,
    pivot_attribute_a,
    pivot_attribute_b,
    attribute_b_date,
    df_test_attribute_a,
    df_desligados_access,
    df_pop_attribute_a,
    testing_attribute_a,
    df_test_attribute_b,
    df_pop_attribute_b,
    testing_attribute_b,
    cap_columns,
    format_df_attribute_a,
    format_df_attribute_b,
    df_summary_atributo_a,
    gera_df_extraction_date,
    gera_df_tipo_teste,
    df_tempestividade,
    gerar_arquivos_output,
    escreve_df_formatado_em_excel,
    escrever_summary,
    caminho_explorer,
    qtd_register,
    gera_df_qtds
)


def attribute_mixed_AB(caminho_input):

    # Files Import

    dfs, BASE_DIR, input_path = files_import(caminho_input)

    df_general_summary     = dfs["df_general_summary"]
    df_system_summary      = dfs["df_system_summary"]
    df_ativos              = dfs["df_ativos"]
    df_desligados          = dfs["df_desligados"]
    df_usuarios_atributo_a = dfs["df_usuarios_atributo_a"]
    df_usuarios_atributo_b = dfs["df_usuarios_atributo_b"]

    
    # Registrando as quantidades originais para validação futura

    qtd_desligados_original = len(df_desligados)
    qtd_ativos_original = len(df_ativos)
    qtd_por_sistema = qtd_register(df_usuarios_atributo_a, df_usuarios_atributo_b)

    # Tratando colunas de ID

    for df in [df_ativos, df_desligados, df_usuarios_atributo_a, df_usuarios_atributo_b]:
        id_normalize_apply(df)



    # Definindo os desligados do FY

    fy_start, fy_end = fy_dates(df_general_summary)

    df_desligados_fy = fy_filter(df_desligados, fy_start, fy_end)

    qtd_desligados_fy = len(df_desligados_fy)



    # Remover falsos positivos de ativos

    df_desligados_fy, qtd_falsos_positivos_ativos = actives_false_positive(df_desligados_fy, df_ativos)



    # Gerar Pivot Tables e DF de datas de Bloqueio

    df_usuarios_atributo_a_pivot   = pivot_attribute_a(df_usuarios_atributo_a)

    df_usuarios_atributo_b_status_pivot = pivot_attribute_b(df_usuarios_atributo_b)

    df_usuarios_atributo_b_data    = attribute_b_date(df_usuarios_atributo_b)



    # Teste do Atributo A

    df_teste_atributo_a = df_test_attribute_a(df_usuarios_atributo_a_pivot, df_usuarios_atributo_b_status_pivot)

    df_desligados_acessos = df_desligados_access(df_desligados_fy, df_teste_atributo_a)

    df_população_teste_atributo_a, sem_acesso, populacao, colunas_sistemas = df_pop_attribute_a(df_desligados_acessos)

    df_população_teste_atributo_a = testing_attribute_a(df_população_teste_atributo_a)



    # Teste do Atributo B

    df_desligados_acessos_b = df_test_attribute_b(df_desligados_fy, df_usuarios_atributo_b_data)

    df_população_teste_atributo_b = df_pop_attribute_b(df_desligados_acessos_b)

    df_população_teste_atributo_b = testing_attribute_b(df_população_teste_atributo_b, df_system_summary)



    # Formatando DFs de Saída

    colunas_para_capitalizar = ["Nome", "Cargo", "Centro de Custo"]
    df_teste_atributo_a = cap_columns(df_população_teste_atributo_a, colunas_para_capitalizar)
    df_teste_atributo_b = cap_columns(df_população_teste_atributo_b, colunas_para_capitalizar)

    df_teste_atributo_a = format_df_attribute_a(df_teste_atributo_a)
    df_teste_atributo_b = format_df_attribute_b(df_teste_atributo_b)



    # DF Summary Attribute A

    df_summary_attribute_a = df_summary_atributo_a(
        qtd_desligados_original=qtd_desligados_original,
        qtd_desligados_fy=qtd_desligados_fy,
        qtd_falsos_positivos_ativos=qtd_falsos_positivos_ativos,
        sem_acesso=sem_acesso,
        populacao=populacao,
        df_populacao_teste=df_população_teste_atributo_a,
        colunas_sistemas=colunas_sistemas
    )



    # DF Extraction Date

    df_extraction_date = gera_df_extraction_date(df_general_summary, df_system_summary, colunas_sistemas)

    # DF Qtds

    df_quantidade = gera_df_qtds(qtd_desligados_original, qtd_ativos_original, qtd_por_sistema, colunas_sistemas)

    # DF Tipo Teste

    df_tipo_teste = gera_df_tipo_teste(df_system_summary, colunas_sistemas)



    # DF Tempestividade

    df_temp_value = df_tempestividade(df_system_summary)



    # Arquivo de Output

    paths = gerar_arquivos_output(df_general_summary, BASE_DIR, input_path)
    caminho_output = paths["caminho_output"]
    nome_cliente = paths["nome_cliente"]

    wb = load_workbook(caminho_output)



    # Escrevendo o excel de teste com formatação

    escreve_df_formatado_em_excel(
        wb = wb,
        df_teste_atributo_a = df_teste_atributo_a,
        df_teste_atributo_b = df_teste_atributo_b
    )



    # Escrevendo o excel de summary com formatação

    escrever_summary(
        wb=wb,
        nome_cliente=nome_cliente,
        fy_start=fy_start,
        df_summary_attribute_a=df_summary_attribute_a,
        df_extraction_date=df_extraction_date,
        df_quantidade=df_quantidade,
        df_tipo_teste=df_tipo_teste,
        df_temp_value=df_temp_value
    )

    wb.save(caminho_output)

    explorer = caminho_explorer(caminho_output)
    
    return explorer
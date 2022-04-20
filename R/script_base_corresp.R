library(dplyr)
library(janitor)
library(readr)
library(readxl)
library(forcats)

# Mudou muito! ---------------------------------------------------------

# Agora os anexos não serão mais retirados dos pdf, mas de Excel que está
# no próprio SCANC (arquivo >>> exportar tabelas para texto >>>
# escolher a tabela >>> desmarcar 'incluir estrutura' >>> vizualizar >>>
# xls >>> esperar um pouco até  abrir  o Excel).

# Caso não existam dados importar seguindo o roteiro abaixo.
# SCANC FCA >>> arquivos >>> importar >>> colocar mês/ano que interessa >>>
# Copiar o caminho abaixo preenchendo os interesses.
# ir em G:\PROGRAMAS UTILIZADOS PELOS AUDITORES\SCANC FCA\Scanc-2\Base_FCA\Versão nova\’ano que interessa’\’mês/ano que interessa’



# 1. Estoque inicial estoque final. ------------------------------------------

# Importando e salvando o arquivo.

tba1q1_06_a_12_2019 <- readxl::read_xlsx("data-raw/excel/tba1q1_06_a_12_2019.xlsx")

saveRDS(tba1q1_06_a_12_2019, "data-raw/rds/tba1q1_06_a_12_2019.rds")


# Função para verificar a igualdade de estoque inicial de um mês
# com o final do mês anterior.
# As sequências 2:7 e 1:6 serão aperfeiçoadas posteriormente
# para 2:n e 1:n-1, quando fizer do ano todo de uma vez n = 12.

# Das quantidades.
inic_final_qtd <- function(cnpj_requerente, combust) {
  temp1 <- tba1q1_06_a_12_2019 |>
  dplyr::filter(CNPJH == cnpj_requerente, PRODUTO == combust)

  temp1[2:7, "QTDGASA1"] == temp1[1:6, "QTDGASA11"]
}

# Das quantidades só de DSM (a função é diferente)
inic_final_qtd_dsm <- function(cnpj_requerente, combust) {
  temp2 <- tba1q1_06_a_12_2019 |>
  dplyr::filter(CNPJH == cnpj_requerente, PRODUTO == combust)

  temp2[2:7, "QTD1"] == temp2[1:6, "QTD11"]
}


# Das bases de cálculo de ST.
inic_final_bcst <- function(cnpj_requerente, combust) {
  temp1 <- tba1q1_06_a_12_2019 |>
  dplyr::filter(CNPJH == cnpj_requerente, PRODUTO == combust)

  temp1[2:7, "VLRBCST1"] == temp1[1:6, "VLRBCST11"]
}

# Exemplos. R. Tudo TRUE!
# Todos estoques iniciais relativos à quantidade e BC ST, de julho até dezembro,
# são iguais ao finais de junho até novembro.

inic_final_qtd(cnpj_requerente = "33337122021396", combust = "GSP")
inic_final_qtd(cnpj_requerente = "33337122021396", combust = "GSL")
inic_final_qtd(cnpj_requerente = "33337122021396", combust = "DSL")
inic_final_qtd(cnpj_requerente = "33337122021396", combust = "S10")
inic_final_qtd_dsm(cnpj_requerente = "33337122021396", combust = "DSM")
inic_final_bcst("33337122021396", "GSP")
inic_final_bcst("33337122021396", "DSM")
inic_final_bcst("33337122021396", "GSL")
inic_final_bcst("33337122021396", "DSL")
inic_final_bcst("33337122021396", "S10")


# 2. Estratégia geral. ----------------------------------------------------

# Na medida em que vamos confrontar as informações do SCANC com as das NFe
# precisammos transformar estas em tibbles.

# Os scripts abaixo leêm todos Anexos I, II, as proporções
# e uma tabela do scanc que dá nome e descrição e código dos combustíveis
# de todos contribuintes de 2019-06 até 2019-12
# Lendo os arquivos baixados e salvando em rds.

# Tabela com as informações sobre nro_nf do Anexo I

tba1q3d_06_a_12_2019 <- readxl::read_xlsx("data-raw/excel/tba1q3d_06_a_12_2019.xlsx")

saveRDS(tba1q3d_06_a_12_2019, "data-raw/rds/an_1_06_a_12_2019.rds")

# Tabela com as proporções.
tba1q2_06_a_12_2019 <- readxl::read_xlsx("data-raw/excel/tba1q2_06_a_12_2019.xlsx")

saveRDS(tba1q2_06_a_12_2019, "data-raw/rds/tba1q2_06_a_12_2019.rds")


# Tabela com as informações sobre nro_nf do Anexo II.
tba2q2d_06_a_12_2019 <- readxl::read_xlsx("data-raw/excel/tba2q2d_06_a_12_2019.xlsx")

saveRDS(tba2q2d_06_a_12_2019, "data-raw/rds/an_2_06_a_12_2019.rds")


# Tabela com os números do produto descrição dos combustíveis.
descri_prod <- readxl::read_xlsx("data-raw/excel/tb_prod_descr_unid.xlsx")
saveRDS(descri_prod, "data-raw/rds/descri_prod.rds")

# Tabela com os códigos de produto da ANP
anp_cod_prod <- readxl::read_xlsx("data-raw/excel/ANP_codigos_de_produto.xlsx", skip = 1) |>
janitor::clean_names()
saveRDS(anp_cod_prod, "data-raw/rds/anp_cod_prod.rds")



# 3. Descompactando os xml. -----------------------------------------------

# Lidas e salvas estas informações em tibbles, vamos agora descompactar os xml das NFe
# e transfomá-los em tibbles.

# Função para descompactar.

descomp_nfe <- function(ano_mes) {

  # Descompactando o arquivo escolhido na pasta ano_mes.
  # O nome desta pasta será composto pela palavra dias e
  # o ano mês do arquivo compactado: por exemplo dias_2018_11.

  utils::unzip(paste0("data-raw/extraidos/", ano_mes, ".zip"),
    exdir = paste("data-raw/zips_dias/", ano_mes)
  )

  # Listando todos que estão na pasta ano_mes.
  diario <- list.files(
    path = paste("data-raw/zips_dias/", ano_mes),
    pattern = "*.zip",
    full.names = TRUE
  )


  descomp <- function(n) {
    utils::unzip(diario[n], exdir = paste("data-raw/xml/", ano_mes))
  }
  # Agora sim, descompactando todos.

  purrr::map(1:length(diario), descomp)

  # Tirando os arquivos que contêm 'EVENTO'.
  # Nomeando os arquivos que não quero.

  arquivos_que_nao_quero <- list.files(paste("data-raw/xml/", ano_mes),
    full.names = TRUE, pattern = "EVENTO"
  )

  # Removendo os arquivos que não quero.
  file.remove(arquivos_que_nao_quero)
}

# Exemplo.
descomp_nfe(ano_mes = "2019_09")

# Cada conjunto de xml foi descompactado numa pasta referente ao seu período.
# Agora é convertê-los numa só tibble: nfe_ipiranga.


# 4. Transformando os xml em tibble, empilhando todas numa só e salvando. -----------

aa <- scancxnfe::nfe_to_tibble(arquivos = NULL, diretorio = "data-raw/xml/ 2019_06/")
bb <- scancxnfe::nfe_to_tibble(arquivos = NULL, diretorio = "data-raw/xml/ 2019_07/")
cc <- scancxnfe::nfe_to_tibble(arquivos = NULL, diretorio = "data-raw/xml/ 2019_08/")
dd <- scancxnfe::nfe_to_tibble(arquivos = NULL, diretorio = "data-raw/xml/ 2019_09/")

# Empilhando e salvando.

nfe_ipiranga <- dplyr::bind_rows(aa, bb, cc, dd)
saveRDS(nfe_ipiranga, "data-raw/rds/nfe_ipiranga.rds")


# 5. Função que faz a correspondência do Anexo I com as NFe. ----------------------------------

verif_corresp_an_1_nfe <- function(an_1, nfe_requerente, vetor_cnpj_refirarias,
                                   periodo, cnpj_solicitante, vetor_nome_prod_scanc) {
  an_1 |>
  # Fazendo o left-join para juntar as duas tabelas.
  dplyr::left_join(nfe_requerente, by = c("NUMNF" = "nNF")) |>
  # Outro left-join para acrescentar a coluna de descrição de produto.
  dplyr::left_join(descri_prod, by = c("PRODNF" = "PRODUTO")) |>
  # Mais um left-join para anexar a tabela com os códigos de produto da ANP.
  # dplyr::left_join(anp_cod_prod, by = c("cProdANP" = "codigo")) |>
  # Criando as colunas que farão parte da tibble arrumada.
  # No momento a "PETROLEO BRASILEIRO S.A" é a única refinaria.
  dplyr::mutate(com_x_sem = dplyr::if_else(
    CNPJ_emit %in% vetor_cnpj_refirarias,
    "COM", "SEM"
  )) |>
  # Criando as colunas de status.
  dplyr::mutate(
    status_quant =
      dplyr::case_when(
        QTD < 0 ~ "devolução",
        abs(QTD - qCom) < 1 ~ "OK",
        TRUE ~ "não corresponde"
      ),
    status_bc_st =
      dplyr::case_when(
        com_x_sem == "SEM" ~ "OK",
        abs(VLRBCST - vBCST) < 1 ~ "OK",
        TRUE ~ "não corresponde"
      )
  ) |>
  # Selecionando, renomeando e ordenando as colunas.
  dplyr::select(
    mes_ano = MESANO,
    cnpj_requerente = CNPJH,
    NUMNF, com_x_sem,
    nome_prod_nfe = xProd,
    #sub_subgrupo_anp = sub_subgrupo,
    #nome_prod_anp = produto,
    descricao_scanc = DESCRICAO,
    nome_prod_scanc = PRODUTO,
    quant_scanc = QTD,
    quant_nfe = qCom,
    status_quant,
    bc_st_scanc = VLRBCST,
    bc_st_nfe = vBCST,
    status_bc_st,
    chave_nfe = chNFe
  ) |>
  dplyr::distinct() |>
  dplyr::filter(
    mes_ano == periodo,
    cnpj_requerente == cnpj_solicitante,
    nome_prod_scanc %in% vetor_nome_prod_scanc
  )
}

# Exemplo.

result_ipiranga_an_1_072019 <- verif_corresp_an_1_nfe(
  an_1 = an_1_06_a_12_2019,
  nfe_requerente = nfe_ipiranga,
  vetor_cnpj_refirarias = c("33000167008862"),
  periodo = "072019",
  cnpj_solicitante = "33337122021396",
  vetor_nome_prod_scanc = c("DSL", "DSM", "GSL", "GSP", "S10")
)

# Salvando.
saveRDS(result_ipiranga_an_1_072019, "data-raw/rds/result_ipiranga_an_1_072019.rds")
openxlsx::write.xlsx(result_ipiranga_an_1_072019, "data-raw/excel/result_ipiranga_an_1_072019.xlsx")


# 6. Proporções (quantidade) Com X SEM. -----------------------------------

# Faxinando e preparando as entradas retiradas do scanc.
entr_scanc <- tba1q2_06_a_12_2019 |>
  dplyr::filter(CNPJH == "33337122021396", MESANO == "072019") |>
  dplyr::mutate(com_x_sem = dplyr::if_else(TIPOICMS == "1", "COM", "SEM")) |>
  dplyr::select(PRODUTO, com_x_sem, QTDENT) |>
  dplyr::rename(qtd_scanc = QTDENT) |>
  dplyr::arrange(PRODUTO, com_x_sem) |>
  tidyr::unite("prod_com_sem", 1:2, sep = "-")

# Faxinando e preparando entradas retiradas das nfe.
entr_nfe <- result_ipiranga_an_1_072019 |>
  dplyr::group_by(nome_prod_scanc, com_x_sem) |>
  dplyr::summarise(tot_quant = sum(quant_nfe)) |>
  dplyr::ungroup() |>
  dplyr::rename(qtd_nfe = tot_quant) |>
  dplyr::arrange(nome_prod_scanc, com_x_sem) |>
  tidyr::unite("prod_com_sem", 1:2, sep = "-")

# Juntando as duas e calculando as diferenças percentuais.
# Diferenças menores que 1% podem ser desprezadas.

result_prop <- entr_scanc |>
  dplyr::full_join(entr_nfe, by = "prod_com_sem") |>
  #dplyr::mutate(dif_perc = round(abs(qtd_scanc - qtd_nfe)*100/qtd_scanc,2)) |>
  dplyr::mutate(variacao = dif_perc(qtd_scanc, qtd_nfe))


# 7. Função que faz a correspondência do Anexo II com as NFe. ----------------------------------

# Repetindo a mesma estratégia e raciocínio para o Anexo II.

verif_corresp_an_2_nfe <- function(an_2, nfe_requerente, periodo, cnpj_solicitante,
                                   vetor_uf_dest, vetor_combs) {
  an_2 |>
  # Juntando com as nfe.
  dplyr::left_join(nfe_requerente, by = c("NUMNF" = "nNF")) |>
  # Juntando com a tabela do scanc com os códigos de produto e suas descrições.
  dplyr::left_join(descri_prod, by = c("PRODNF" = "PRODUTO")) |>
  # Juntando com a tabela da ANP com códigos, nome de produto e sub-subgrupo;
  # dplyr::left_join(anp_cod_prod, by = c("cProdANP" = "codigo")) |>
  # Fazendo as correções para gasolinas.
  dplyr::mutate(
    icms_dest_nfe =
      dplyr::case_when(
        DESCRICAO == "GASOLINA A PREMIUM,62Y" ~ vICMSSTDest / 0.75,
        DESCRICAO == "GASOLINA A" ~ vICMSSTDest / 0.73,
        TRUE ~ vICMSSTDest
      )
  ) |>
  # Criando as colunas de status.
  dplyr::mutate(
    status_quant =
      dplyr::case_when(
        QTD < 0 ~ "devolução",
        abs(QTD - qCom) < 1 ~ "OK",
        TRUE ~ "não corresponde"
      ),
    status_icms_dev_dest =
      dplyr::case_when(
        VLRICMSST < 0 ~ "devolução",
        abs(VLRICMSST - icms_dest_nfe) < 1 ~ "OK",
        TRUE ~ "não corresponde"
      )
  ) |>
  # Selecionando, ordenando e renomeando as colunas.
  dplyr::select(
    mes_ano = MESANO,
    cnpj_requerente = CNPJH,
    NUMNF,
    uf_dest = UF_dest,
    #sub_subgrupo_anp = sub_subgrupo,
    #nome_prod_anp = produto,
    cod_prod_scanc = PRODNF,
    descri_prod_scanc = DESCRICAO,
    nome_prod_scanc = PRODUTO,
    nome_prod_nfe = xProd,
    qtd_scanc = QTD,
    qtd_nfe = qCom,
    status_quant,
    icms_dest_scanc = VLRICMSST,
    icms_dest_nfe,
    status_icms_dev_dest,
    chave_nfe = chNFe
  ) |>
  dplyr::filter(
    mes_ano == periodo,
    cnpj_requerente == cnpj_solicitante,
    uf_dest %in% vetor_uf_dest,
    nome_prod_scanc %in% vetor_combs
  ) |>
  dplyr::distinct()
}

# Exemplo.

result_ipiranga_an_2_072019 <- verif_corresp_an_2_nfe(
  an_2 = an_2_06_a_12_2019,
  nfe_requerente = nfe_ipiranga,
  periodo = "072019",
  cnpj_solicitante = "33337122021396",
  vetor_uf_dest = c("MG", "ES", "SP"),
  vetor_combs = c("DSL", "DSM", "GSL", "GSP", "S10")
)



# Salvando em rds e xlsx.
saveRDS(result_ipiranga_an_2_072019, "data-raw/rds/result_ipiranga_an_2_072019.rds")
openxlsx::write.xlsx(result_ipiranga_an_2_072019, "data-raw/excel/result_ipiranga_an_2_072019.xlsx")




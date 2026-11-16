###############################################################################
# POF 2017-2018 (IBGE): leitura e agregação dos dados
#   - Resposta composicional: os 11 grupos de despesa de consumo da POF, em R$, sem reagrupar
#   - Saída: uma linha por unidade de consumo, com o gasto mensal (R$) em cada grupo,
#     o gasto total e as covariáveis. As proporções NÃO são calculadas aqui.
#     (arquivo saida_pof/pof_composicional.csv)
#
# Unidade de análise: unidade de consumo (UC) = COD_UPA + NUM_DOM + NUM_UC.
# É a unidade em que a POF registra as despesas (na maioria dos domicílios,
# 1 domicílio = 1 UC).
#
# COMO USAR
# 1. Baixe do IBGE (página dos microdados da POF 2017-2018) o zip dos "Dados"
#    e o zip da "Documentação". Descompacte tudo dentro da pasta PASTA_POF.
#    O script procura os arquivos por nome, em qualquer subpasta.
# 2. Rode o script de cima para baixo. O CSV final vai para PASTA_SAIDA.
#
# ATENÇÃO: o script não foi testado com os arquivos reais. Os pontos mais
# prováveis de ajuste são a leitura do dicionário (abas/colunas do .xls) e a
# detecção do nível "grupo de despesa" no tradutor. Há mensagens de erro e
# diagnósticos nesses trechos.
###############################################################################

# 0. PACOTES --------------------------------------------------------------------
pacotes <- c("readr", "readxl", "dplyr", "tidyr", "stringr")
novos <- setdiff(pacotes, rownames(installed.packages()))
if (length(novos)) install.packages(novos)
invisible(lapply(pacotes, library, character.only = TRUE))

# 1. CONFIGURAÇÃO ------------------------------------------------------------------
PASTA_POF    <- "raw_data"          # pasta onde você descompactou dados + documentação
PASTA_SAIDA  <- "data"
DIVISOR_COD  <- 100            # V9001 tem 7 dígitos; o tradutor usa 5 (V9001 %/% 100)
N_AMOSTRA    <- NULL           # ex.: 15000 para testar rápido; NULL = base toda
set.seed(2026)
dir.create(PASTA_SAIDA, showWarnings = FALSE)

# 2. FUNÇÕES AUXILIARES ---------------------------------------------------------------
norm <- function(x) {                       # minúsculas, sem acento, sem espaços nas pontas
  tolower(stringi::stri_trans_general(trimws(as.character(x)), "Latin-ASCII"))
}

achar <- function(padrao) {                 # procura arquivo pelo nome dentro de PASTA_POF
  r <- list.files(PASTA_POF, pattern = padrao, recursive = TRUE,
                  ignore.case = TRUE, full.names = TRUE)
  if (!length(r)) stop("Arquivo não encontrado (padrão '", padrao, "') em ", PASTA_POF)
  r[1]
}

ARQ_DICIONARIO <- achar("dicion.*\\.xlsx?$")
ARQ_TRADUTOR   <- tryCatch(
  {
    cand <- list.files(PASTA_POF, pattern = "despesa.*\\.xlsx?$", recursive = TRUE,
                       ignore.case = TRUE, full.names = TRUE)
    cand <- cand[grepl("tradutor|geral", cand, ignore.case = TRUE)]
    if (!length(cand)) stop("sem tradutor")
    message("Candidatos a tradutor de despesa: ", paste(basename(cand), collapse = " | "))
    cand[order(!grepl("geral", cand, ignore.case = TRUE))][1]
  },
  error = function(e) stop(
    "O arquivo Tradutor_Despesa_Geral.xls não está na pasta '", PASTA_POF, "'.\n",
    "Ele não vem no zip de Dados nem no de Documentação: procure na página de microdados ",
    "da POF 2017-2018 o arquivo de 'Tradutores das tabelas', descompacte-o dentro de '",
    PASTA_POF, "' e rode de novo.", call. = FALSE))
message("Dicionário: ", ARQ_DICIONARIO)
message("Tradutor:   ", ARQ_TRADUTOR)

# Lê uma aba de Excel achando a linha de cabeçalho (a que contém `padrao`)
ler_excel_cabecalho <- function(arq, aba, padrao) {
  raw <- read_excel(arq, sheet = aba, col_names = FALSE, col_types = "text",
                    .name_repair = "minimal")
  lin <- which(apply(raw, 1, function(r) any(grepl(padrao, norm(r)))))[1]
  if (is.na(lin)) stop("Cabeçalho '", padrao, "' não encontrado na aba ", aba)
  nm <- norm(unname(unlist(raw[lin, ])))
  vazio <- is.na(nm) | nm == ""
  nm[vazio] <- paste0("col", which(vazio))
  out <- raw[-seq_len(lin), ]
  names(out) <- make.unique(nm)
  out
}

# Dicionário de variáveis de um registro (posição, tamanho, decimais, nome)
ler_dicionario <- function(nome, aba = NULL) {
  abas <- excel_sheets(ARQ_DICIONARIO)
  if (is.null(aba)) {
    chave <- gsub("_", " ", tolower(nome))
    cand  <- abas[grepl(chave, norm(abas), fixed = TRUE)]
    if (!length(cand))
      stop("Sem aba para ", nome, " no dicionário. Abas disponíveis: ",
           paste(abas, collapse = " | "), "\nUse o argumento `aba`.")
    aba <- cand[which.min(nchar(cand))]
  }
  d  <- ler_excel_cabecalho(ARQ_DICIONARIO, aba, "posicao")
  nm <- names(d)
  i_ini <- grep("posicao", nm)[1];  i_tam <- grep("^tamanho", nm)[1]
  i_dec <- grep("^decimais", nm)[1]; i_var <- grep("codigo", nm)[1]
  tibble(
    ini = suppressWarnings(as.integer(d[[i_ini]])),
    tam = suppressWarnings(as.integer(d[[i_tam]])),
    dec = if (is.na(i_dec)) 0L else suppressWarnings(as.integer(d[[i_dec]])),
    var = trimws(d[[i_var]])
  ) %>%
    mutate(dec = coalesce(dec, 0L)) %>%
    filter(!is.na(ini), !is.na(tam), !is.na(var), var != "")
}

# Lê um registro de largura fixa (só as variáveis pedidas) e converte p/ numérico
ler_pof <- function(nome, vars, opcionais = character(), aba = NULL) {
  arq <- achar(paste0("^", nome, "\\.txt$"))
  dic <- ler_dicionario(nome, aba)
  falt <- setdiff(vars, dic$var)
  if (length(falt))
    stop("Variáveis ausentes no dicionário de ", nome, ": ", paste(falt, collapse = ", "),
         "\nVariáveis disponíveis: ", paste(head(dic$var, 60), collapse = ", "))
  usar <- dic %>% filter(var %in% c(vars, opcionais)) %>%
    distinct(var, .keep_all = TRUE) %>% arrange(ini)
  pos <- fwf_positions(usar$ini, usar$ini + usar$tam - 1, usar$var)
  d <- read_fwf(arq, pos, col_types = cols(.default = col_character()),
                trim_ws = TRUE, progress = FALSE)
  for (i in seq_len(nrow(usar))) {
    v <- usar$var[i]; x <- d[[v]]
    num <- suppressWarnings(as.numeric(x))
    # decimais implícitos (sem ponto no arquivo) -> divide por 10^dec
    if (usar$dec[i] > 0 && !any(grepl(".", x, fixed = TRUE))) num <- num / 10^usar$dec[i]
    d[[v]] <- num
  }
  message(nome, ": ", format(nrow(d), big.mark = "."), " linhas")
  d
}

# 3. TRADUTOR: código do item -> categoria de despesa --------------------------------
tr <- ler_excel_cabecalho(ARQ_TRADUTOR, 1, "^codigo")
col_cod <- names(tr)[grep("^codigo", names(tr))][1]
cols_txt <- setdiff(names(tr), col_cod)

# Os 11 grupos de despesa de consumo da POF
GRUPOS_REGEX <- paste0("^[0-9 .-]*", c(
  "alimentacao", "habitacao", "vestuario", "transporte", "higiene",
  "assistencia a saude", "educacao", "recreacao", "servicos pessoais",
  "fumo", "despesas diversas"))
# Nome de cada grupo, na mesma ordem de GRUPOS_REGEX (viram as colunas gasto_<nome>)
GRUPOS_NOMES <- c("alimentacao", "habitacao", "vestuario", "transporte", "higiene",
                  "saude", "educacao", "recreacao", "servicos_pessoais",
                  "fumo", "despesas_diversas")

# Detecta automaticamente qual coluna do tradutor traz o nível "grupo"
pontua <- sapply(cols_txt, function(cn) {
  v <- unique(norm(tr[[cn]]))
  sum(sapply(GRUPOS_REGEX, function(p) any(grepl(p, v))))
})
col_grupo <- names(which.max(pontua))
message("Coluna do tradutor usada como 'grupo': ", col_grupo,
        " (", max(pontua), " de 11 grupos reconhecidos)")
if (max(pontua) < 5)
  stop("Não reconheci os grupos de despesa no tradutor. Veja `names(tr)` e `head(tr)`.")

trad <- tr %>%
  transmute(cod5 = suppressWarnings(as.integer(.data[[col_cod]])),
            g = norm(.data[[col_grupo]])) %>%
  filter(!is.na(cod5)) %>%
  distinct(cod5, .keep_all = TRUE) %>%
  mutate(idx = vapply(g, function(x) {
    h <- which(str_detect(x, GRUPOS_REGEX))
    if (length(h)) h[1] else NA_integer_
  }, integer(1)),
  cat = GRUPOS_NOMES[idx]) %>%
  filter(!is.na(cat))   # itens fora dos 11 grupos de consumo (impostos etc.) ficam de fora
print(count(trad, cat))

# 4. DESPESAS POR UC -------------------------------------------------------------------
# valor mensal = V8000_DEFLA (deflacionado) * FATOR_ANUALIZACAO / 12
agrega_despesa <- function(nome) {
  ler_pof(nome, c("COD_UPA", "NUM_DOM", "NUM_UC", "V9001",
                  "V8000_DEFLA", "FATOR_ANUALIZACAO")) %>%
    transmute(COD_UPA, NUM_DOM, NUM_UC,
              cod5 = V9001 %/% DIVISOR_COD,
              valor = V8000_DEFLA * FATOR_ANUALIZACAO / 12)
}

desp <- bind_rows(
  agrega_despesa("DESPESA_COLETIVA"),
  agrega_despesa("DESPESA_INDIVIDUAL"),
  agrega_despesa("CADERNETA_COLETIVA"),
  agrega_despesa("ALUGUEL_ESTIMADO")
) %>%
  inner_join(trad, by = "cod5")

if (nrow(desp) == 0)
  stop("Nenhum item casou com o tradutor. Teste DIVISOR_COD <- 1 ou confira os códigos.")

CATS <- GRUPOS_NOMES
comp <- desp %>%
  group_by(COD_UPA, NUM_DOM, NUM_UC, cat) %>%
  summarise(v = sum(valor, na.rm = TRUE), .groups = "drop") %>%
  pivot_wider(names_from = cat, values_from = v, values_fill = 0)
for (g in setdiff(CATS, names(comp))) comp[[g]] <- 0   # grupo sem nenhum registro
comp <- comp %>%
  mutate(total = rowSums(across(all_of(CATS)))) %>%
  filter(total > 0) %>%                      # UC sem nenhum gasto de consumo não entra
  rename_with(~ paste0("gasto_", .x), all_of(CATS))   # gasto_<grupo> em R$/mês

# 5. COVARIÁVEIS ---------------------------------------------------------------------
# Todas as variáveis e códigos de categoria abaixo foram conferidos no
# "Dicionários de váriaveis.xls" da POF 2017-2018. O script PARA com erro se aparecer
# um código que não está no dicionário ou se uma chave estiver duplicada.
chave_uc  <- c("COD_UPA", "NUM_DOM", "NUM_UC")
chave_dom <- c("COD_UPA", "NUM_DOM")

checa_cat <- function(x, validos, nome) {
  fora <- !is.na(x) & !(x %in% validos)
  if (any(fora))
    stop("Variável ", nome, " tem código fora do dicionário: ",
         paste(head(unique(x[fora])), collapse = ", "))
  invisible(TRUE)
}
uc1 <- function(x) { x <- x[!is.na(x)]; if (length(x)) x[1] else NA_real_ }   # 1º valor não vazio
bin <- function(x, ok) ifelse(is.na(x), NA_integer_, as.integer(x %in% ok))     # 0/1 preservando NA

UF_SIGLA <- c("11" = "RO", "12" = "AC", "13" = "AM", "14" = "RR", "15" = "PA", "16" = "AP",
              "17" = "TO", "21" = "MA", "22" = "PI", "23" = "CE", "24" = "RN", "25" = "PB",
              "26" = "PE", "27" = "AL", "28" = "SE", "29" = "BA", "31" = "MG", "32" = "ES",
              "33" = "RJ", "35" = "SP", "41" = "PR", "42" = "SC", "43" = "RS", "50" = "MS",
              "51" = "MT", "52" = "GO", "53" = "DF")

## 5.1 MORADOR (pessoas) -> agregado por UC -------------------------------------------
mor <- ler_pof("MORADOR",
               c("UF", "TIPO_SITUACAO_REG", "ESTRATO_POF", "COD_UPA", "NUM_DOM", "NUM_UC",
                 "V0306", "V0403", "V0404", "V0405", "V0406", "V0415", "V0418",
                 "ANOS_ESTUDO", "NIVEL_INSTRUCAO", "PESO_FINAL",
                 "RENDA_TOTAL", "RENDA_DISP_PC", "RENDA_MONET_PC",
                 "RENDA_NAO_MONET_PC", "DEDUCAO_PC"))

checa_cat(mor$UF, as.integer(names(UF_SIGLA)), "UF")
checa_cat(mor$TIPO_SITUACAO_REG, 1:2, "TIPO_SITUACAO_REG")
checa_cat(mor$V0306, 1:19, "V0306")
checa_cat(mor$V0404, 1:2, "V0404")
checa_cat(mor$V0405, c(1:5, 9), "V0405")
checa_cat(mor$V0406, 1:2, "V0406")
checa_cat(mor$V0415, 1:2, "V0415")
checa_cat(mor$V0418, 1:2, "V0418")
checa_cat(mor$NIVEL_INSTRUCAO, 1:7, "NIVEL_INSTRUCAO")

# Variáveis que o dicionário define no nível da UC precisam ser constantes dentro da UC
for (v in c("UF", "TIPO_SITUACAO_REG", "ESTRATO_POF", "PESO_FINAL", "RENDA_TOTAL",
            "RENDA_DISP_PC", "RENDA_MONET_PC", "RENDA_NAO_MONET_PC", "DEDUCAO_PC")) {
  k <- mor %>% group_by(across(all_of(chave_uc))) %>%
    summarise(k = n_distinct(.data[[v]], na.rm = TRUE), .groups = "drop")
  if (any(k$k > 1)) stop(v, " varia dentro da UC; não é uma variável de UC. Me envie esta mensagem.")
}

uc <- mor %>%
  group_by(across(all_of(chave_uc))) %>%
  summarise(
    n_ref            = sum(V0306 == 1),                      # deve ser exatamente 1 por UC
    n_morad          = n(),                                  # todas as linhas de moradores da UC
    n_criancas       = sum(V0403 < 15, na.rm = TRUE),
    n_idosos         = sum(V0403 >= 65, na.rm = TRUE),
    n_empregados_dom = sum(V0306 %in% c(18, 19)),            # empregado(a) doméstico(a) e parentes
    n_plano_saude    = sum(V0406 == 1, na.rm = TRUE),
    n_estudantes     = sum(V0415 == 1, na.rm = TRUE),
    n_estud_privada  = sum(V0415 == 1 & V0418 == 1, na.rm = TRUE),
    n_estud_publica  = sum(V0415 == 1 & V0418 == 2, na.rm = TRUE),
    uf               = uc1(UF),
    situacao         = uc1(TIPO_SITUACAO_REG),
    estrato_pof      = uc1(ESTRATO_POF),
    peso_final       = uc1(PESO_FINAL),
    renda_total        = uc1(RENDA_TOTAL),
    renda_disp_pc      = uc1(RENDA_DISP_PC),
    renda_monet_pc     = uc1(RENDA_MONET_PC),
    renda_nao_monet_pc = uc1(RENDA_NAO_MONET_PC),
    deducao_pc         = uc1(DEDUCAO_PC),
    idade_ref        = V0403[V0306 == 1][1],                 # pessoa de referência da UC
    sexo_ref         = V0404[V0306 == 1][1],
    cor_ref          = V0405[V0306 == 1][1],
    anos_estudo_ref  = ANOS_ESTUDO[V0306 == 1][1],
    nivel_instrucao_ref = NIVEL_INSTRUCAO[V0306 == 1][1],
    .groups = "drop")

message("UCs sem pessoa de referência única (excluídas): ", sum(uc$n_ref != 1))
uc <- uc %>% filter(n_ref == 1) %>% select(-n_ref)

## 5.2 DOMICILIO (chave: COD_UPA + NUM_DOM; não tem NUM_UC) ----------------------------
dom <- ler_pof("DOMICILIO",
               c("COD_UPA", "NUM_DOM", "V0201", "V0205", "V0206", "V0207", "V02111",
                 "V0212", "V0213", "V0217", "V0220", "V6199"))
if (any(duplicated(dom[chave_dom]))) stop("DOMICILIO tem chave COD_UPA+NUM_DOM duplicada.")
checa_cat(dom$V0201, 1:3, "V0201");  checa_cat(dom$V0207, 1:6, "V0207")
checa_cat(dom$V0212, 1:5, "V0212");  checa_cat(dom$V0213, 1:6, "V0213")
checa_cat(dom$V0217, 1:7, "V0217");  checa_cat(dom$V0220, 1:2, "V0220")
checa_cat(dom$V6199, 1:4, "V6199")

dom <- dom %>% transmute(
  COD_UPA, NUM_DOM,
  tipo_domicilio = factor(V0201, levels = 1:3,
                          labels = c("Casa", "Apartamento", "Cômodo/cortiço")),
  n_comodos     = V0205,
  n_dormitorios = V0206,
  n_banheiros   = V02111,                       # banheiros de uso exclusivo dos moradores
  agua_rede_geral = bin(V0207, 1),              # 1 = rede geral de distribuição
  esgoto_rede     = bin(V0212, 1),              # 1 = rede geral/pluvial/fossa ligada à rede; NA = não se aplica
  lixo_coletado   = bin(V0213, 1:2),            # 1 = coletado (direto ou em caçamba)
  rua_pavimentada = bin(V0220, 1),
  condicao_ocupacao = factor(V0217, levels = 1:7, labels = c(
    "Próprio - já pago", "Próprio - ainda pagando", "Alugado", "Cedido por empregador",
    "Cedido por familiar", "Cedido de outra forma", "Outra condição")),
  inseguranca_alimentar = factor(V6199, levels = 1:4, labels = c(
    "Segurança", "Insegurança leve", "Insegurança moderada", "Insegurança grave"),
    ordered = TRUE))

## 5.3 CONDICOES_VIDA (chave: UC) ---------------------------------------------------------
cv <- ler_pof("CONDICOES_VIDA", c(chave_uc, "V6101"), aba = "Condições de Vida")
if (any(duplicated(cv[chave_uc])))
  stop("CONDICOES_VIDA tem mais de uma linha por UC. Me envie esta mensagem.")
checa_cat(cv$V6101, 1:6, "V6101")
cv <- cv %>% transmute(COD_UPA, NUM_DOM, NUM_UC,
                       dificuldade_fim_mes = V6101)   # 1 = muita dificuldade ... 6 = muita facilidade

## 5.4 MORADOR_QUALI_VIDA (variáveis derivadas pelo IBGE, nível UC) -------------------
qv <- ler_pof("MORADOR_QUALI_VIDA", c(chave_uc, "C5", "C6", "C7"),
              aba = "Morador - Qualidade de Vida")
checa_cat(qv$C5, 1:6, "C5"); checa_cat(qv$C6, 1:6, "C6"); checa_cat(qv$C7, 1:10, "C7")
for (v in c("C5", "C6", "C7")) {
  k <- qv %>% group_by(across(all_of(chave_uc))) %>%
    summarise(k = n_distinct(.data[[v]], na.rm = TRUE), .groups = "drop")
  if (any(k$k > 1)) stop(v, " varia dentro da UC; não é uma variável de UC. Me envie esta mensagem.")
}
qv <- qv %>%
  group_by(across(all_of(chave_uc))) %>%
  summarise(ocupacao_ref = uc1(C5), composicao_familiar = uc1(C6),
            decil_renda = uc1(C7), .groups = "drop") %>%
  mutate(
    ocupacao_ref = factor(ocupacao_ref, levels = 1:6, labels = c(
      "Empregado doméstico", "Empregado privado", "Militar e empregado do setor público",
      "Conta própria", "Empregador", "Fora da força de trabalho e outros casos")),
    composicao_familiar = factor(composicao_familiar, levels = 1:6, labels = c(
      "Um adulto sem criança", "Um adulto com ao menos uma criança",
      "Mais de um adulto sem criança", "Mais de um adulto com ao menos uma criança",
      "Um ou mais idosos, com ou sem crianças",
      "Um ou mais idosos, com ao menos um adulto, com ou sem crianças")),
    decil_renda = as.integer(decil_renda))   # 1 = 1º décimo (mais pobre) ... 10 = 10º

## 5.5 INVENTARIO (bens duráveis; códigos conferidos no Cadastro de Produtos, quadro 14) --
ARQ_CADASTRO <- achar("^cadastro.de.produtos\\.xlsx?$")
cad <- read_excel(ARQ_CADASTRO, col_types = "text")
names(cad) <- c("quadro", "cod", "desc")
cad14 <- cad %>% filter(quadro == "14") %>%
  mutate(cod = as.numeric(cod), desc_n = norm(desc))

BENS <- c(automovel = 1403001, motocicleta = 1403101, bicicleta = 1403201,
          computador = 1401901, tablet = 1402001, ar_condicionado = 1402101,
          maquina_lavar_roupas = 1401201, maquina_lavar_loucas = 1400801, freezer = 1400201)
BENS_KW <- c(automovel = "automovel", motocicleta = "motocicleta", bicicleta = "bicicleta",
             computador = "microcomputador", tablet = "tablet", ar_condicionado = "ar-condicionado",
             maquina_lavar_roupas = "lavar roupas", maquina_lavar_loucas = "lavar loucas",
             freezer = "freezer")
for (b in names(BENS)) {          # o código precisa ter a descrição esperada no cadastro
  d <- cad14$desc_n[cad14$cod == BENS[[b]]]
  if (length(d) != 1 || !grepl(BENS_KW[[b]], d, fixed = TRUE))
    stop("Código ", BENS[[b]], " não corresponde a '", b, "' no Cadastro de Produtos.")
}

inv <- ler_pof("INVENTARIO", c(chave_uc, "V9001", "V9005"))
if (!all(inv$V9001 %in% cad14$cod)) stop("INVENTARIO tem código de bem fora do quadro 14.")
if (any(is.na(inv$V9005) | inv$V9005 < 1)) stop("INVENTARIO: quantidade (V9005) fora de 1 a 34.")

inv_tipos <- inv %>% group_by(across(all_of(chave_uc))) %>%
  summarise(n_tipos_bens = n_distinct(V9001), .groups = "drop")   # tipos de bens (de 32 possíveis)
inv_bens <- inv %>%
  filter(V9001 %in% BENS) %>%
  mutate(bem = paste0("n_", names(BENS)[match(V9001, BENS)])) %>%
  group_by(across(all_of(chave_uc)), bem) %>%
  summarise(q = sum(V9005), .groups = "drop") %>%                  # quantidade existente
  pivot_wider(names_from = bem, values_from = q, values_fill = 0)
BEM_COLS <- paste0("n_", names(BENS))
for (b in setdiff(BEM_COLS, names(inv_bens))) inv_bens[[b]] <- 0

## 5.6 RENDIMENTOS por fonte (R$/mês da UC; valor deflacionado, bruto) ----------------------
ARQ_TRAD_REND <- achar("^tradutor.rendimento\\.xlsx?$")
trr <- read_excel(ARQ_TRAD_REND, col_types = "text") %>%
  transmute(cod5 = suppressWarnings(as.integer(Codigo)),
            n2 = suppressWarnings(as.integer(Nivel_2)),
            n3 = suppressWarnings(as.integer(Nivel_3))) %>%
  filter(!is.na(cod5)) %>%                     # tira as linhas "Ver arquivo ..." (não monetário e variação patrimonial)
  mutate(fonte = case_when(
    n3 == 111 ~ "rend_trab_empregado",   n3 == 112 ~ "rend_trab_empregador",
    n3 == 113 ~ "rend_trab_conta_propria",
    n3 == 121 ~ "rend_transf_inss",      n3 == 122 ~ "rend_transf_prev_publica",
    n3 == 123 ~ "rend_transf_prev_privada", n3 == 124 ~ "rend_transf_prog_sociais",
    n3 == 125 ~ "rend_transf_pensao_doacao", n3 == 126 ~ "rend_transf_outras",
    n2 == 13  ~ "rend_aluguel",          n2 == 14  ~ "rend_outras_rendas",
    TRUE ~ NA_character_))
if (anyNA(trr$fonte) || any(duplicated(trr$cod5))) stop("Tradutor_Rendimento: código sem fonte ou duplicado.")
REND_COLS <- c("rend_trab_empregado", "rend_trab_empregador", "rend_trab_conta_propria",
               "rend_transf_inss", "rend_transf_prev_publica", "rend_transf_prev_privada",
               "rend_transf_prog_sociais", "rend_transf_pensao_doacao", "rend_transf_outras",
               "rend_aluguel", "rend_outras_rendas")
stopifnot(setequal(unique(trr$fonte), REND_COLS))

agrega_rend <- function(nome, aba) {
  ler_pof(nome, c(chave_uc, "V9001", "V8500_DEFLA", "FATOR_ANUALIZACAO"), aba = aba) %>%
    transmute(COD_UPA, NUM_DOM, NUM_UC, cod5 = V9001 %/% 100,
              valor = V8500_DEFLA * FATOR_ANUALIZACAO / 12)    # mesmo critério das despesas
}
rend_raw <- bind_rows(agrega_rend("RENDIMENTO_TRABALHO", "Rendimento do Trabalho"),
                      agrega_rend("OUTROS_RENDIMENTOS", "Outros Rendimentos"))
message("Linhas de rendimento fora do tradutor (movimentação financeira, trabalho não ",
        "remunerado etc.; não entram): ", sum(!rend_raw$cod5 %in% trr$cod5), " de ", nrow(rend_raw))
rend <- rend_raw %>%
  inner_join(trr[, c("cod5", "fonte")], by = "cod5") %>%
  group_by(across(all_of(chave_uc)), fonte) %>%
  summarise(v = sum(valor, na.rm = TRUE), .groups = "drop") %>%
  pivot_wider(names_from = fonte, values_from = v, values_fill = 0)
for (b in setdiff(REND_COLS, names(rend))) rend[[b]] <- 0
rend$rend_monetario_total <- rowSums(rend[, REND_COLS])

## 5.7 ESTRATOS (tabela "Estratos POF 2017-2018": faixas de ESTRATO_POF por tipo de área) --
ESTRATOS <- data.frame(
  ini  = c(1101, 1103, 1108, 1201, 1202, 1203, 1301, 1307, 1308, 1311, 1401, 1403, 1404, 1501, 1504, 1506, 1512, 1601, 1603, 1604, 1605, 1701, 1702, 1706, 2101, 2104, 2105, 2114, 2201, 2204, 2210, 2301, 2307, 2310, 2321, 2401, 2403, 2404, 2409, 2501, 2504, 2506, 2512, 2601, 2604, 2607, 2616, 2701, 2704, 2705, 2709, 2801, 2803, 2804, 2807, 2901, 2907, 2910, 2926, 3101, 3107, 3110, 3131, 3201, 3203, 3206, 3212, 3301, 3310, 3319, 3331, 3501, 3510, 3516, 3537, 4101, 4106, 4109, 4125, 4201, 4203, 4205, 4218, 4301, 4307, 4310, 4325, 5001, 5004, 5010, 5101, 5103, 5104, 5113, 5201, 5204, 5207, 5218, 5301, 5307),
  fim  = c(1102, 1107, 1111, 1201, 1202, 1204, 1306, 1307, 1310, 1316, 1402, 1403, 1405, 1503, 1505, 1511, 1519, 1602, 1603, 1604, 1607, 1701, 1705, 1708, 2103, 2104, 2113, 2125, 2203, 2209, 2217, 2306, 2309, 2320, 2330, 2402, 2403, 2408, 2412, 2503, 2505, 2511, 2518, 2603, 2606, 2615, 2624, 2703, 2704, 2708, 2713, 2802, 2803, 2806, 2810, 2906, 2909, 2925, 2942, 3106, 3109, 3130, 3149, 3202, 3205, 3211, 3215, 3309, 3318, 3330, 3337, 3509, 3515, 3536, 3553, 4105, 4108, 4124, 4135, 4202, 4204, 4217, 4226, 4306, 4309, 4324, 4335, 5003, 5009, 5013, 5102, 5103, 5112, 5118, 5203, 5206, 5217, 5225, 5306, 5308),
  area = c("capital", "resto_uf", "rural", "capital", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "resto_rm", "resto_uf", "rural", "capital", "rural"),
  stringsAsFactors = FALSE)

area_estrato_fun <- function(e) {
  idx <- vapply(e, function(x) {
    if (is.na(x)) return(NA_integer_)
    h <- which(x >= ESTRATOS$ini & x <= ESTRATOS$fim)
    if (length(h) == 1) h else NA_integer_
  }, integer(1))
  ESTRATOS$area[idx]
}

## 5.8 MONTAGEM -------------------------------------------------------------------------------
perdidas <- anti_join(comp, uc, by = chave_uc)
message("UCs com despesa que não casaram com MORADOR: ", nrow(perdidas))

dados <- comp %>%
  inner_join(uc, by = chave_uc) %>%
  mutate(
    renda_pc = renda_total / n_morad,   # renda total da UC (inclui variação patrimonial) / moradores
    regiao   = factor(uf %/% 10, levels = 1:5,
                      labels = c("Norte", "Nordeste", "Sudeste", "Sul", "Centro-Oeste")),
    regiao   = relevel(regiao, ref = "Sudeste"),
    uf_sigla = factor(unname(UF_SIGLA[as.character(uf)]), levels = unname(UF_SIGLA)),
    situacao = factor(situacao, levels = 1:2, labels = c("Urbano", "Rural")),
    sexo_ref = factor(sexo_ref, levels = 1:2, labels = c("Homem", "Mulher")),
    cor_ref  = factor(cor_ref, levels = c(1, 2, 3, 4, 5, 9),
                      labels = c("Branca", "Preta", "Amarela", "Parda", "Indígena",
                                 "Sem declaração")),
    nivel_instrucao_ref = factor(nivel_instrucao_ref, levels = 1:7, labels = c(
      "Sem instrução", "Fundamental incompleto", "Fundamental completo",
      "Médio incompleto", "Médio completo", "Superior incompleto", "Superior completo"),
      ordered = TRUE))

n_base <- nrow(dados)
dados <- dados %>%
  left_join(dom, by = chave_dom) %>%
  left_join(cv,  by = chave_uc) %>%
  left_join(qv,  by = chave_uc) %>%
  left_join(inv_tipos, by = chave_uc) %>%
  left_join(inv_bens,  by = chave_uc) %>%
  left_join(rend, by = chave_uc)
stopifnot(nrow(dados) == n_base)         # os cruzamentos não podem criar nem perder linhas

# Tipo de área a partir do estrato (todos os estratos precisam estar na tabela)
area_cod <- area_estrato_fun(dados$estrato_pof)
if (anyNA(area_cod)) stop("ESTRATO_POF fora da tabela de estratos: ",
                          paste(head(unique(dados$estrato_pof[is.na(area_cod)])), collapse = ", "))
dados$area_estrato <- factor(area_cod, levels = c("capital", "resto_rm", "resto_uf", "rural"),
                             labels = c("Município da capital", "Resto da RM",
                                        "Resto da UF (urbano)", "Rural"))

# Bens duráveis: NA só se a UC não tem nenhum registro de inventário; senão ausência = 0
dados <- dados %>%
  mutate(across(all_of(BEM_COLS), ~ if_else(is.na(n_tipos_bens), NA_real_, coalesce(.x, 0))),
         across(all_of(c(REND_COLS, "rend_monetario_total")), ~ coalesce(.x, 0)))

G <- paste0("gasto_", CATS)              # colunas de gasto (R$/mês)
dados <- dados %>%
  select(COD_UPA, NUM_DOM, NUM_UC, estrato_pof, peso_final,
         total, all_of(G),
         renda_total, renda_pc, renda_disp_pc, renda_monet_pc, renda_nao_monet_pc, deducao_pc,
         n_morad, n_criancas, n_idosos, n_empregados_dom,
         n_plano_saude, n_estudantes, n_estud_privada, n_estud_publica,
         idade_ref, sexo_ref, cor_ref, anos_estudo_ref, nivel_instrucao_ref,
         uf_sigla, regiao, situacao,
         tipo_domicilio, n_comodos, n_dormitorios, n_banheiros, agua_rede_geral,
         esgoto_rede, lixo_coletado, rua_pavimentada, condicao_ocupacao,
         inseguranca_alimentar, dificuldade_fim_mes,
         area_estrato, ocupacao_ref, composicao_familiar, decil_renda,
         n_tipos_bens, all_of(BEM_COLS),
         all_of(REND_COLS), rend_monetario_total)

# Linhas sem informação nas variáveis centrais são excluídas; nas demais, NA é mantido
# (vários NA são legítimos: "Não aplicável" no dicionário).
central <- c("total", G, "peso_final", "renda_pc", "n_morad", "idade_ref", "sexo_ref",
             "cor_ref", "anos_estudo_ref", "situacao", "regiao")
antes <- nrow(dados)
dados <- dados[complete.cases(dados[, central]), ]
message("Excluídas por NA em variável central: ", antes - nrow(dados))

if (!is.null(N_AMOSTRA)) dados <- slice_sample(dados, n = min(N_AMOSTRA, nrow(dados)))

dados <- as.data.frame(dados)
write_csv(dados, file.path(PASTA_SAIDA, "pof_composicional.csv"))
message("Base final: ", nrow(dados), " UCs, ", ncol(dados), " colunas")

# 6. CHECAGENS ---------------------------------------------------------------------------------
cat("\n== NA por coluna (só as que têm) ==\n")
na_col <- colSums(is.na(dados)); print(na_col[na_col > 0])

# (a) Chaves únicas
stopifnot(!any(duplicated(dados[chave_uc])))

# (b) Gastos somam o total
stopifnot(max(abs(rowSums(dados[, G]) - dados$total)) < 1e-6)

# (c) Comparação com os números oficiais do IBGE (POF 2017-2018, Brasil, valores POND.):
#     despesa de consumo média mensal familiar ~ R$ 3.764,51;
#     participação nas despesas de consumo: habitação 36,6%, transporte 18,1%, alimentação 17,5%.
#     Os números do IBGE usam PESO_FINAL e razão de totais, então comparamos da mesma forma.
w <- dados$peso_final
cat("\n== Validação ponderada (esperado pelo IBGE entre parênteses) ==\n")
cat("Despesa de consumo média mensal (R$ 3.764,51): ",
    round(sum(dados$total * w) / sum(w), 2), "\n")
part <- colSums(dados[, G] * w) / sum(dados$total * w)
cat("Habitação (0,366): ",   round(part["gasto_habitacao"], 3), "\n")
cat("Transporte (0,181): ",  round(part["gasto_transporte"], 3), "\n")
cat("Alimentação (0,175): ", round(part["gasto_alimentacao"], 3), "\n")

# (d) Renda: nossa renda_pc contra a soma das rendas per capita oficiais (devem andar juntas)
ok <- complete.cases(dados$renda_pc, dados$renda_monet_pc, dados$renda_nao_monet_pc)
cat("\nCorrelação renda_pc x (renda_monet_pc + renda_nao_monet_pc): ",
    round(cor(dados$renda_pc[ok], (dados$renda_monet_pc + dados$renda_nao_monet_pc)[ok]), 3), "\n")
cat("Participação dos grupos no gasto (não ponderada):\n")
print(round(colSums(dados[, G]) / sum(dados$total), 3))

# (e) Estrato x situação do domicílio: duas fontes independentes, devem coincidir
cat("\n== Tipo de área (estrato) x situação do domicílio ==\n")
print(table(dados$area_estrato, dados$situacao))
n_disc <- sum((dados$area_estrato == "Rural") != (dados$situacao == "Rural"))
if (n_disc > 0) warning(n_disc, " UCs com estrato rural/urbano diferente da situação do domicílio.")

# (f) Rendimentos: soma das fontes x renda monetária do IBGE (RENDA_MONET_PC * moradores),
#     só UCs sem empregados domésticos (o IBGE os exclui da renda per capita)
sem_dom <- dados$n_empregados_dom == 0 & !is.na(dados$renda_monet_pc) & dados$renda_monet_pc > 0
razao <- dados$rend_monetario_total[sem_dom] / (dados$renda_monet_pc[sem_dom] * dados$n_morad[sem_dom])
cat("\n== Soma das fontes de rendimento / renda monetária do IBGE ==\n")
print(round(quantile(razao, c(.01, .05, .25, .5, .75, .95, .99), na.rm = TRUE), 3))
cat("Proporção de UCs com razão entre 0,99 e 1,01: ",
    round(mean(razao > 0.99 & razao < 1.01, na.rm = TRUE), 3), "\n")


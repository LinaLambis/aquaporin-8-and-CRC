###############################################################################
# 0. Load the data
###############################################################################

library(ggplot2)
library(readxl)
library(dplyr)
library(tableone)
library(writexl)
library(writexl)
library(officer)
library(flextable)
library(rstatix)
library(tibble)
library(FSA)


###############################################################################
# 1. Load the data
###############################################################################

data <- read_excel("~/Desktop/AQP8/data/Base de Datos _ Resultados AQP8 Corregida 10092025.xlsx")
data_wb = read_excel("~/Desktop/AQP8/data/WesternBlotAQP8.xlsx")


###############################################################################
# 2. Depuration
###############################################################################

data <- data %>%
  mutate(
    Sexo = factor(recode(Sexo,
                         `1` = "Male",
                         `2` = "Female")),
    
    Raza = factor(recode(Raza,
                         `1` = "Indigenous",
                         `2` = "Black",
                         `3` = "White",
                         `4` = "Other")),
    
    Procedencia = factor(recode(Procedencia,
                                `1` = "Cartagena",
                                `5` = "Arjona",
                                `6` = "Arroyo Hondo",
                                `8` = "Calamar",
                                `11` = "Córdoba",
                                `12` = "Clemencia",
                                `13` = "El Carmen de Bolívar",
                                `17` = "Magangué",
                                `18` = "Mahates",
                                `19` = "Margarita",
                                `20` = "María La Baja",
                                `28` = "San Estanislao",
                                `30` = "San Jacinto",
                                `32` = "San Juan Nepomuceno",
                                `39` = "Soplaviento",
                                `41` = "Tiquisio",
                                `42` = "Turbaco",
                                `45` = "Zambrano",
                                `46` = "Other Department")),
    
    Residencia = factor(recode(Residencia,
                               `1` = "Urban",
                               `2` = "Rural")),
    
    Afiliación = factor(recode(Afiliación,
                               `1` = "Contributive",
                               `2` = "Subsidized",
                               `4` = "Linked")),
    
    TiempoTto = factor(recode(TiempoTto,
                              `1` = "<3 months",
                              `2` = "3–6 months",
                              `3` = ">6 months"),
                       levels = c("<3 months", "3–6 months", ">6 months"), ordered = TRUE),
    
    IMC = factor(recode(IMC,
                        `1` = "<20",
                        `2` = "20–25",
                        `3` = "26–30",
                        `4` = "30–35",
                        `5` = ">35"),
                 levels = c("<20", "20–25", "26–30", "30–35", ">35"), ordered = TRUE),
    
    Sintom = factor(recode(Sintom,
                           `1` = "Intestinal obstruction",
                           `2` = "Acute abdominal pain",
                           `3` = "Bowel habit changes",
                           `4` = "Rectal bleeding",
                           `5` = "Anemia",
                           `6` = "Other")),
    
    Tamizaje = factor(recode(Tamizaje,
                             `1` = "Fecal occult blood test",
                             `3` = "Colonoscopy",
                             `4` = "No screening")),
    
    CCRfam = factor(recode(CCRfam,
                           `1` = "No",
                           `2` = "Yes - Sporadic",
                           `3` = "Yes - FAP")),
    
    Colono = factor(recode(Colono,
                           `1` = "Stenosing lesion",
                           `2` = "Exophytic lesion",
                           `3` = "Single polyp",
                           `4` = "Multiple polyps",
                           `5` = "Ulcerative lesion",
                           `7` = "Generalized colitis",
                           `8` = "Not described")),
    
    Localiz = factor(recode(Localiz,
                            `1` = "Appendix",
                            `2` = "Cecum",
                            `3` = "Right colon",
                            `4` = "Transverse colon",
                            `5` = "Left colon",
                            `6` = "Sigmoid colon",
                            `7` = "Rectum",
                            `9` = "Not reported")),
    
    ACE = factor(recode(ACE,
                        `1` = "Not available",
                        `2` = "<5",
                        `3` = ">5"),
                 levels = c("Not available", "<5", ">5"), ordered = TRUE),
    
    Histol = factor(recode(Histol,
                           `1` = "Adenocarcinoma NOS",
                           `2` = "Mucinous adenocarcinoma",
                           `4` = "Signet ring cell adenocarcinoma",
                           `7` = "Neuroendocrine carcinoma",
                           `9` = "Other")),
    
    GradoHist = factor(recode(GradoHist,
                              `1` = "Well differentiated",
                              `2` = "Moderately differentiated",
                              `3` = "Poorly differentiated"),
                       levels = c("Well differentiated", "Moderately differentiated", "Poorly differentiated"), ordered = TRUE),
    
    Angio = factor(recode(Angio, `1` = "No", `2` = "Yes")),
    Linfo = factor(recode(Linfo, `1` = "No", `2` = "Yes")),
    IHQ = factor(recode(IHQ, `1` = "No", `2` = "Yes")),
    
    T = factor(recode(T,
                      `1` = "Tx",
                      `2` = "T0",
                      `3` = "T1",
                      `4` = "T2",
                      `5` = "T3",
                      `6` = "T4a",
                      `7` = "T4b",
                      `8` = "No data"),
               levels = c("Tx", "T0", "T1", "T2", "T3", "T4a", "T4b", "No data"), ordered = TRUE),
    
    N = factor(recode(N,
                      `1` = "Nx",
                      `2` = "N0",
                      `3` = "N1a",
                      `4` = "N1b",
                      `5` = "N1c",
                      `6` = "N2",
                      `7` = "N2a",
                      `8` = "N2b"),
               levels = c("Nx", "N0", "N1a", "N1b", "N1c", "N2", "N2a", "N2b"), ordered = TRUE),
    
    M = factor(recode(M,
                      `1` = "M0",
                      `2` = "M1a",
                      `3` = "M1b",
                      `4` = "No data"),
               levels = c("M0", "M1a", "M1b", "No data"), ordered = TRUE),
    
    SiteM = factor(recode(SiteM,
                          `1` = "Liver",
                          `2` = "Lung",
                          `4` = "Brain",
                          `5` = "No metastasis")),
    
    EC = factor(recode(EC,
                       `1` = "0",
                       `2` = "I",
                       `3` = "IIA",
                       `4` = "IIB",
                       `5` = "IIC",
                       `6` = "IIIA",
                       `7` = "IIIB",
                       `8` = "IIIC",
                       `9` = "IVA",
                       `10` = "IVB"),
                levels = c("0", "I", "IIA", "IIB", "IIC", "IIIA", "IIIB", "IIIC", "IVA", "IVB"), ordered = TRUE),
    
    QT = factor(recode(QT,
                       `1` = "No",
                       `2` = "Neoadjuvant",
                       `3` = "Adjuvant",
                       `4` = "Not described")),
    
    RT = factor(recode(RT,
                       `1` = "No",
                       `2` = "Neoadjuvant",
                       `3` = "Adjuvant",
                       `4` = "Not described")),
    
    Terabiol = factor(recode(Terabiol, `1` = "No", `2` = "Yes")),
    
    Progres = factor(recode(Progres, `1` = "No", `2` = "Yes")),
    
    ECcodi = factor(recode(ECcodi,
                           `1` = "Local/Early",
                           `2` = "Regional/Advanced",
                           `3` = "Metastatic/Advanced"),
                    levels = c("Local/Early", "Regional/Advanced", "Metastatic/Advanced"), ordered = TRUE),
    
    NEDAD = factor(recode(NEDAD,
                          `1` = "<=35",
                          `2` = "36–48",
                          `3` = "49–61",
                          `4` = "62–74",
                          `5` = "75–87",
                          `6` = "88+"),
                   levels = c("<=35", "36–48", "49–61", "62–74", "75–87", "88+"), ordered = TRUE),
    
    NUEVAEDAD = factor(recode(NUEVAEDAD,
                              `1` = "<47",
                              `2` = "48–74",
                              `3` = "75–99"),
                       levels = c("<47", "48–74", "75–99"), ordered = TRUE),
    
    Subtipoanatóm = factor(recode(Subtipoanatóm,
                                  `1` = "Proximal colon",
                                  `2` = "Distal colon",
                                  `3` = "Rectum")),
    
    Edaddos = factor(recode(Edaddos,
                            `1` = "<=50",
                            `2` = "51+"),
                     levels = c("<=50", "51+"), ordered = TRUE),
    
    TNM = factor(recode(TNM,
                        `1` = "0",
                        `2` = "I",
                        `3` = "IIA",
                        `4` = "IIB",
                        `5` = "IIC",
                        `6` = "IIIA",
                        `7` = "IIIB",
                        `8` = "IIIC",
                        `9` = "IVA",
                        `10` = "IVB"),
                 levels = c("0", "I", "IIA", "IIB", "IIC", "IIIA", "IIIB", "IIIC", "IVA", "IVB"), ordered = TRUE),
    
    TNM2 = factor(recode(TNM2,
                         `1` = "Early",
                         `2` = "Advanced"),
                  levels = c("Early", "Advanced"), ordered = TRUE),
    
    TipoSintoma = factor(recode(TipoSintoma,
                                `1` = "Acute",
                                `2` = "Chronic")),
    
    TNMR = factor(recode(TNMR,
                         `1` = "0",
                         `2` = "I",
                         `3` = "II",
                         `4` = "III",
                         `5` = "IV"),
                  levels = c("0", "I", "II", "III", "IV"), ordered = TRUE),
    
    DiametroCod = factor(recode(DiametroCod,
                                `1` = ">5 cm",
                                `2` = "<5 cm"),
                         levels = c("<5 cm", ">5 cm"), ordered = TRUE),
    
    EspesorCod = factor(recode(EspesorCod,
                               `1` = ">5 cm",
                               `2` = "<5 cm"),
                        levels = c("<5 cm", ">5 cm"), ordered = TRUE)
  )


# Add a metastasis y/n column
data$metastasis <- ifelse(data$SiteM == "No metastasis", "No", "Yes")

# Convertir la variable metastasis en factor para asegurarnos del orden
data$metastasis <- factor(data$metastasis, levels = c("No", "Yes"))


# Add a column with histologic differentiation binomial
data$Histol_dif <- ifelse(data$GradoHist %in% c("Well differentiated"), "Well differentiated", 
                          "Some grade of undifferentiation")

# Convertir la variable Histol_dif en factor para asegurarnos del orden
data$Histol_dif <- factor(data$Histol_dif, levels = c("Well differentiated", "Some grade of undifferentiation"))

#Eliminate rows number 103 - 108
data <- data[-c(103:108), ]

# Eliminate some NAs columns
data$...48 = NULL
data$...49 = NULL

# Create group variable combining location and stage
data$group <- paste(data$Subtipoanatóm, data$TNM2, sep = " - ")

# Order group factor levels
data$group <- factor(data$group, levels = c(
  "Proximal colon - Early",
  "Proximal colon - Advanced",
  "Distal colon - Early",
  "Distal colon - Advanced"
))



################################
# 2.1 Data WB
################################

data_wb <- data_wb %>%
  mutate(
    ESTADIO = recode(
      ESTADIO,
      "AVANZADO" = "Advanced",
      "TEMPRANO" = "Early",
      "CONTROL"  = "Control"
    )
  )




###############################################################################
# 3. Table 1
###############################################################################

###### Version #1 General -----------------------------------------------------

# Crear tabla general (sin estratificación)
table1 <- CreateTableOne(
  vars = c("Sexo", "Residencia", "Sintom", "Colono",
           "Histol", "GradoHist", "Espesor", "Diametro",
           "TNMR", "ACE", "TNM2", "Progres", "metastasis",
           "Raza", "IMC", "Localiz", "Edad"),  # Se añadió TNMR en lugar de TNM, y se incluyen Localiz y Edad directamente
  data = data,
  factorVars = c("Sexo", "Residencia", "Sintom", "Colono",
                 "Histol", "GradoHist", "TNMR", "ACE", "TNM2",
                 "Progres", "metastasis", "Raza", "IMC", "Localiz"),
  test = FALSE  # No se necesitan pruebas estadísticas para una tabla general
)

# Convertir a data.frame con nombres de fila
table1_df <- as.data.frame(print(table1, quote = FALSE, noSpaces = TRUE, printToggle = FALSE))
table1_df <- rownames_to_column(table1_df, var = "Variable")

# Crear flextable
ft <- flextable(table1_df) %>%
  autofit() %>%
  fontsize(size = 10, part = "all") %>%
  set_table_properties(width = .95, layout = "autofit")

# Exportar a Word
doc <- read_docx() %>%
  body_add_par("Table 1. General Clinicopathological Characteristics", style = "heading 1") %>%
  body_add_par("Note: This table presents the overall characteristics of the study population.", style = "Normal") %>%
  body_add_flextable(ft)

print(doc, target = "~/Desktop/AQP8/Table1_General_Characteristics.docx")









####### Tabla S1 -------------------------------------------------------------

# 1) Prepare data with Localizacion_tmp and create stage group from TNMR
data2 <- data %>%
  mutate(
    Localizacion_tmp = case_when(
      Localiz %in% c("Cecum", "Right colon", "Transverse colon") ~ "Proximal colon",
      Localiz %in% c("Left colon", "Sigmoid colon")             ~ "Distal colon",
      TRUE ~ NA_character_
    ),
    Edad_tmp = ifelse(Edad < 50, "<50 years", "≥50 years"),
    TNMR = factor(TNMR, levels = c("0", "I", "II", "III", "IV")),
    # Create early/advanced grouping based on TNMR
    Stage_Group = case_when(
      TNMR %in% c("0", "I", "II") ~ "Early",
      TNMR %in% c("III", "IV")    ~ "Advanced",
      TRUE                        ~ NA_character_
    ),
    Stage_Group = factor(Stage_Group, levels = c("Early", "Advanced"))
  )

# 2) Create a combined stratification variable for clearer presentation
data2 <- data2 %>%
  mutate(
    Location_Stage = interaction(Localizacion_tmp, Stage_Group, sep = " - ")
  )

# 3) Build TableOne with stratification by the combined variable
table1 <- CreateTableOne(
  vars       = c("Sexo", "Residencia", "Sintom", "Colono",
                 "Histol", "GradoHist", "Espesor", "Diametro",
                 "TNMR", "ACE", "Progres", "metastasis",
                 "Raza", "IMC"),
  strata     = "Location_Stage",
  data       = data2,
  factorVars = c("Sexo", "Residencia", "Sintom", "Colono",
                 "Histol", "GradoHist", "TNMR", "ACE", 
                 "Progres", "metastasis", "Raza", "IMC"),
  test       = TRUE
)

# 4) Convert to data.frame and wrap in flextable
table1_df <- as.data.frame(print(table1,
                                 quote = FALSE,
                                 noSpaces = TRUE,
                                 printToggle = FALSE)) %>%
  rownames_to_column("Variable")

ft <- flextable(table1_df) %>%
  autofit() %>%
  fontsize(size = 10, part = "all") %>%
  set_table_properties(width = 1, layout = "autofit")

# 5) Export to Word with updated title/note
doc <- read_docx() %>%
  body_add_par(
    "Table 1. Clinicopathological characteristics by tumor location and TNM stage",
    style = "heading 1"
  ) %>%
  body_add_par(
    "Note: Stratified by tumor location (proximal vs distal colon) and TNM stage (early: stages 0-II; advanced: stages III-IV). P-values were calculated using chi-squared tests for categorical variables and Kruskal–Wallis for continuous variables.",
    style = "Normal"
  ) %>%
  body_add_flextable(ft)

print(doc, target = "~/Desktop/AQP8/Table1_ClinicopathCharacteristics_by_Location_TNMStage.docx")








###############################################################################
# Incorporating both a multi-level "Localiz" and a two-level "Localiz_tmp"
###############################################################################

library(dplyr)
library(flextable)
library(officer)
library(tibble)
library(rlang)

# Esta función compara la expresión de AQP8 (ΔΔCT) entre diferentes niveles
# de una variable categórica, realizando pruebas estadísticas apropiadas
# según normalidad (Shapiro-Wilk).
compare_AQP8 <- function(data, varname, expression_col = "ΔΔCT AQP8") {
  expr_sym <- rlang::sym(expression_col)
  df <- data
  
  # Redefinimos la codificación para "Localiz_tmp"
  # y para "Localiz" (todas las categorías)
  if (varname == "Localiz_tmp") {
    df <- df %>%
      filter(!is.na(Localiz), !is.na(!!expr_sym)) %>%
      mutate(group = case_when(
        Localiz %in% c("Cecum", "Right colon", "Transverse colon") ~ "Proximal",
        Localiz %in% c("Left colon", "Sigmoid colon")             ~ "Distal",
        TRUE ~ NA_character_
      ))
  } else if (varname == "Localiz") {
    # Se toman todas las categorías presentes en 'Localiz'
    df <- df %>%
      filter(!is.na(Localiz), !is.na(!!expr_sym)) %>%
      mutate(group = Localiz)
  } else if (varname == "Edad") {
    df <- df %>%
      filter(!is.na(Edad), !is.na(!!expr_sym)) %>%
      mutate(group = ifelse(Edad < 50, "<50", "≥50"))
  } else if (varname == "TNM") {
    df <- df %>%
      filter(!is.na(TNM), !is.na(!!expr_sym)) %>%
      mutate(group = case_when(
        TNM %in% c("0", "I", "IIA", "IIB", "IIC") ~ "Early",
        TNM %in% c("IIIA", "IIIB", "IIIC", "IVA", "IVB") ~ "Advanced",
        TRUE ~ NA_character_
      ))
  } else {
    df <- df %>%
      filter(!is.na(.data[[varname]]), !is.na(!!expr_sym)) %>%
      mutate(group = .data[[varname]])
  }
  
  # Eliminamos filas sin "group"
  df <- df %>% filter(!is.na(group))
  
  # Calculamos estadísticos descriptivos
  summary_stats <- df %>%
    group_by(group) %>%
    summarise(
      Mean = mean(!!expr_sym, na.rm = TRUE),
      SD   = sd(!!expr_sym, na.rm = TRUE),
      n    = n(),
      .groups = "drop"
    ) %>%
    mutate(`Mean ± SD` = sprintf("%.2f ± %.2f", Mean, SD)) %>%
    select(Group = group, `Mean ± SD`)
  
  # Evaluamos normalidad por grupo
  shapiro_p <- df %>%
    group_by(group) %>%
    summarise(p = shapiro.test(!!expr_sym)$p.value, .groups = "drop") %>%
    pull(p)
  is_normal <- all(shapiro_p > 0.05)
  
  # Determinamos prueba estadística en función del número de grupos
  n_groups <- n_distinct(df$group)
  if (n_groups == 2) {
    test <- if (is_normal) {
      t.test(df[[expression_col]] ~ df$group)
    } else {
      wilcox.test(df[[expression_col]] ~ df$group)
    }
  } else {
    test <- if (is_normal) {
      summary(aov(df[[expression_col]] ~ df$group))[[1]][["Pr(>F)"]][1]
    } else {
      kruskal.test(df[[expression_col]] ~ df$group)$p.value
    }
  }
  
  # Extraemos p-valor
  p_val <- if (is.list(test)) signif(test$p.value, 3) else signif(test, 3)
  
  # Etiquetamos la prueba
  test_name <- if (n_groups == 2) {
    if (is_normal) "t-test" else "Mann–Whitney U"
  } else {
    if (is_normal) "ANOVA" else "Kruskal–Wallis"
  }
  
  # Añadimos fila con el test y el p-valor
  result <- summary_stats %>%
    add_row(Group = test_name, `Mean ± SD` = paste("p =", p_val))
  
  return(result)
}

# Creamos "Localiz_tmp" en la base:
data2 <- data %>%
  mutate(
    Localiz_tmp = case_when(
      Localiz %in% c("Cecum", "Right colon", "Transverse colon") ~ "Proximal",
      Localiz %in% c("Left colon", "Sigmoid colon")             ~ "Distal",
      TRUE ~ NA_character_
    )
  )

# Seleccionamos las variables de interés, incluidas "Localiz" y "Localiz_tmp"
vars <- c(
  "Edad", "Sexo", "Raza", "Localiz_tmp", "Localiz", "Progres", "GradoHist",
  "TNM", "ECcodi", "IMC", "TNMR", "metastasis"
)

# Ejecutamos la comparación para cada variable seleccionada
results_list <- lapply(vars, function(v) compare_AQP8(data2, v, expression_col = "ΔΔCT AQP8"))
names(results_list) <- vars

# Construimos una tabla consolidada
comparacion_lancet <- tibble()

for (v in names(results_list)) {
  df <- results_list[[v]]
  
  # Extraemos la fila con la prueba
  test_row <- df[nrow(df), ]
  test <- gsub(" p =.*", "", test_row$Group)
  p_val <- gsub(".*p = ", "", test_row$`Mean ± SD`)
  
  # Excluimos la última fila (test)
  grupos <- df[-nrow(df), ]
  
  # Reestructuramos para ver 2 grupos o más de 2
  if (nrow(grupos) == 2) {
    comparacion_lancet <- bind_rows(
      comparacion_lancet,
      tibble(
        Variable = v,
        Group1   = grupos$Group[1],
        `AQP8 1` = grupos$`Mean ± SD`[1],
        Group2   = grupos$Group[2],
        `AQP8 2` = grupos$`Mean ± SD`[2],
        `p-value`= p_val,
        Test     = test
      )
    )
  } else if (nrow(grupos) > 2) {
    comparacion_lancet <- bind_rows(
      comparacion_lancet,
      tibble(
        Variable = v,
        Group1   = paste(grupos$Group, collapse = "\n"),
        `AQP8 1` = paste(grupos$`Mean ± SD`, collapse = "\n"),
        Group2   = NA,
        `AQP8 2` = NA,
        `p-value`= p_val,
        Test     = test
      )
    )
  }
}

# Renombramos las variables en la tabla final
comparacion_lancet <- comparacion_lancet %>%
  mutate(Variable = recode(
    Variable,
    "Edad"        = "Age",
    "Sexo"        = "Sex",
    "Raza"        = "Race",
    "Localiz_tmp" = "Tumor location (Proximal/Distal)",
    "Localiz"     = "Tumor location (5 levels)",
    "Progres"     = "Tumor progression",
    "GradoHist"   = "Histologic grade",
    "TNM"         = "TNM stage",
    "ECcodi"      = "Clinical stage",
    "IMC"         = "BMI",
    "TNMR"        = "TNMR",
    "metastasis"  = "Metastasis"
  ))

# Exportamos la tabla a Word en estilo "lancet"
ft_lancet <- flextable(comparacion_lancet) %>%
  autofit() %>%
  set_table_properties(layout = "autofit") %>%
  fontsize(size = 10, part = "all")

doc <- read_docx() %>%
  body_add_par("Table. Comparison of AQP8 expression (ΔΔCT) across clinical subgroups", style = "heading 1") %>%
  body_add_par(
    "Values represent mean ± standard deviation. P-values calculated using the appropriate test based on normality. Kruskal–Wallis used for >2 groups. AQP8 = aquaporin 8.",
    style = "Normal"
  ) %>%
  body_add_flextable(ft_lancet)

print(doc, target = "~/Desktop/AQP8/Tabla_Comparacion_AQP8_Clinico_Lancet_Localiz_Both.docx")





###############################################################################
# 3. Expression plots
###############################################################################


# Load required libraries
library(ggplot2)
library(dplyr)
library(rstatix)
library(ggpubr)
library(officer)
library(flextable)

# Check normality in each group
normality_test <- data %>%
  group_by(group) %>%
  summarise(
    n = n(),
    shapiro_p = shapiro.test(`ΔΔCT AQP8`)$p.value,
    normal = shapiro_p > 0.05
  )

print(normality_test)

# Perform statistical comparisons between all groups
# If data is normal, use ANOVA + post-hoc Tukey
# Otherwise, use Kruskal-Wallis + post-hoc Dunn
if(all(normality_test$normal)) {
  # Parametric test: ANOVA + Tukey
  global_result <- aov(`ΔΔCT AQP8` ~ group, data = data) %>% summary()
  print("ANOVA Results:")
  print(global_result)
  
  comparisons <- data %>% tukey_hsd(`ΔΔCT AQP8` ~ group) %>%
    select(group1, group2, p.adj, p.adj.signif)
  
  test_used <- "ANOVA with post-hoc Tukey HSD"
} else {
  # Non-parametric test: Kruskal-Wallis + Dunn
  global_result <- kruskal.test(`ΔΔCT AQP8` ~ group, data = data)
  print("Kruskal-Wallis Results:")
  print(global_result)
  
  comparisons <- data %>% dunn_test(`ΔΔCT AQP8` ~ group, p.adjust.method = "bonferroni") %>%
    select(group1, group2, p.adj, p.adj.signif)
  
  test_used <- "Kruskal-Wallis with post-hoc Dunn's test and Bonferroni correction"
}

# Create descriptive statistics dataframe
descriptive_stats <- data %>%
  group_by(group) %>%
  summarise(
    n = n(),
    mean = mean(`ΔΔCT AQP8`, na.rm = TRUE),
    median = median(`ΔΔCT AQP8`, na.rm = TRUE),
    sd = sd(`ΔΔCT AQP8`, na.rm = TRUE),
    se = sd/sqrt(n),
    min = min(`ΔΔCT AQP8`, na.rm = TRUE),
    max = max(`ΔΔCT AQP8`, na.rm = TRUE),
    Q1 = quantile(`ΔΔCT AQP8`, 0.25, na.rm = TRUE),
    Q3 = quantile(`ΔΔCT AQP8`, 0.75, na.rm = TRUE)
  )

# Format descriptive statistics for publication
if(all(normality_test$normal)) {
  # For parametric data: mean (SD)
  formatted_stats <- descriptive_stats %>%
    mutate(
      `Group` = group,
      `N` = n,
      `Expression (mean ± SD)` = sprintf("%.3f ± %.3f", mean, sd),
      `Range [min, max]` = sprintf("[%.3f, %.3f]", min, max)
    ) %>%
    select(`Group`, `N`, `Expression (mean ± SD)`, `Range [min, max]`)
} else {
  # For non-parametric data: median [IQR]
  formatted_stats <- descriptive_stats %>%
    mutate(
      `Group` = group,
      `N` = n,
      `Expression (median [IQR])` = sprintf("%.3f [%.3f, %.3f]", median, Q1, Q3),
      `Range [min, max]` = sprintf("[%.3f, %.3f]", min, max)
    ) %>%
    select(`Group`, `N`, `Expression (median [IQR])`, `Range [min, max]`)
}

# Format post-hoc comparison results
formatted_comparisons <- comparisons %>%
  mutate(
    `Group 1` = group1,
    `Group 2` = group2,
    `Adjusted p-value` = sprintf("%.3f", p.adj),
    `Significance` = case_when(
      p.adj < 0.001 ~ "p < 0.001",
      p.adj < 0.01 ~ "p < 0.01",
      p.adj < 0.05 ~ "p < 0.05",
      TRUE ~ "NS"
    )
  ) %>%
  select(`Group 1`, `Group 2`, `Adjusted p-value`, `Significance`)

# Simplified Word document creation
# Create a basic Word document without complex formatting
doc <- read_docx()

# Add first table - descriptive statistics
doc <- body_add_par(doc, "Table 1. AQP8 Expression Levels Across Colon Location and TNM Stage Groups", style = "heading 1")
doc <- body_add_par(doc, "")

# Create a simple flextable for descriptive statistics
stats_table <- flextable(formatted_stats)
stats_table <- autofit(stats_table)
# Add simple caption
stats_table <- set_caption(stats_table, "Descriptive statistics of AQP8 expression levels (ΔΔCT) across different colon location and TNM stage groups.")

# Add the table to document
doc <- body_add_flextable(doc, stats_table)
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "")

# Add second table - statistical comparisons
doc <- body_add_par(doc, "Table 2. Statistical Comparisons Between Groups", style = "heading 1")
doc <- body_add_par(doc, "")

# Create a simple flextable for comparisons
comp_table <- flextable(formatted_comparisons)
comp_table <- autofit(comp_table)
# Add caption
comp_table <- set_caption(comp_table, paste("Multiple comparisons of AQP8 expression levels between groups using", test_used, "."))

# Add the table to document
doc <- body_add_flextable(doc, comp_table)
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "NS: not significant.")
doc <- body_add_par(doc, "All p-values were adjusted for multiple comparisons using the Bonferroni method.")

# Save Word document
print(doc, target = "~/Desktop/AQP8/AQP8_expression_results_for_Lancet.docx")

# Create pastel color palette suitable for The Lancet
pastel_colors <- c("#C6DBEF", "#9ECAE1", "#FDCDAC", "#FDD0A2")

# Create minimalist violin plot with individual points
ggplot(data, aes(x = group, y = `2^ΔΔCT AQP8`, fill = group)) +
  # Violin plots with black borders, no trimming
  geom_violin(trim = FALSE, color = "black", linetype = "solid", 
              size = 0.8, alpha = 0.5, scale = "width", width = 0.7) +
  
  # Individual data points with jitter - larger, solid fill, thicker border
  geom_jitter(width = 0.12, alpha = 0.65, size = 1.8, shape = 21, 
              color = "black", stroke = 0.6, aes(fill = group)) +
  
  # Small boxplot
  geom_boxplot(
    width = 0.1, 
    outlier.shape = NA, 
    color = "black",
    fill = "white",
    alpha = 1,
    size = 0.6
  ) +
  
  # Logarithmic Y-axis
  scale_y_log10() +
  
  # Labels
  labs(
    x = "",
    y = expression(AQP8~expression~({Delta*Delta*C[t]}))
  ) +
  
  # Pastel colors
  scale_fill_manual(values = pastel_colors) +
  
  # Minimalist theme with only left and bottom margins
  theme_minimal(base_size = 12) +
  theme(
    axis.title.y = element_text(face = "bold", size = 14, margin = margin(r = 15)),
    axis.text = element_text(size = 11, color = "black"),
    axis.text.x = element_text(angle = 45, hjust = 1),
    panel.grid = element_blank(),
    
    # Keep only left and bottom axis lines, make them thicker
    axis.line.x.top = element_blank(),
    axis.line.y.right = element_blank(),
    axis.line.x.bottom = element_line(color = "black", size = 0.8),
    axis.line.y.left = element_line(color = "black", size = 0.8),
    
    # Remove border
    panel.border = element_blank(),
    
    # Keep only left and bottom ticks
    axis.ticks.x.top = element_blank(),
    axis.ticks.y.right = element_blank(),
    axis.ticks.x.bottom = element_line(color = "black", size = 0.8),
    axis.ticks.y.left = element_line(color = "black", size = 0.8),
    
    legend.position = "none",
    plot.margin = margin(20, 20, 20, 20)
  )









##### AQP8 vs. Progression -----------------------------------------------------

# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `Progres`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Progres), ]

# Convertir la variable Progres en factor para asegurarnos del orden
data_filtrada$Progres <- factor(data_filtrada$Progres, levels = c("No", "Yes"))

# Crear el gráfico
grafico_violines_progres <- ggplot(data_filtrada, aes(x = Progres, y = `ΔΔCT AQP8`, fill = Progres)) +
  geom_violin(trim = FALSE, color = "black", linetype = "solid", size = 0.6, alpha = 0.7) +
  geom_boxplot(
    width = 0.1, 
    outlier.shape = 21, 
    outlier.size = 2, 
    outlier.stroke = 0.5, 
    outlier.fill = "white", 
    color = "black",
    fill = "white"
  ) +
  labs(
    x = "Progression",
    y = expression(AQP8~(2^{-Delta*Delta*C[t]}))
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    axis.title.y = element_text(face = "bold", size = 14),
    axis.text = element_text(size = 12, color = "black"),
    panel.grid = element_blank(),
    axis.line = element_line(color = "black"),
    axis.ticks = element_line(color = "black"),
    legend.position = "none"
  ) +
  scale_fill_manual(values = c("No" = "#9ECAE1", "Yes" = "#FDD0A2"))

# Mostrar el gráfico
print(grafico_violines_progres)





##### AQP8 vs. TNMR -----------------------------------------------------

# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `TNMR`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$TNMR), ]

# Crear el gráfico
grafico_violines_TNMR <- ggplot(data_filtrada, aes(x = TNMR, y = `ΔΔCT AQP8`, fill = TNMR)) +
  geom_violin(trim = FALSE, color = "black", linetype = "solid", size = 0.6, alpha = 0.7) +
  geom_boxplot(
    width = 0.1, 
    outlier.shape = 21, 
    outlier.size = 2, 
    outlier.stroke = 0.5, 
    outlier.fill = "white", 
    color = "black",
    fill = "white"
  ) +
  labs(
    x = "TNM",
    y = expression(AQP8~(2^{-Delta*Delta*C[t]}))
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    axis.title.y = element_text(face = "bold", size = 14),
    axis.text = element_text(size = 12, color = "black"),
    panel.grid = element_blank(),
    axis.line = element_line(color = "black"),
    axis.ticks = element_line(color = "black"),
    legend.position = "none"
  ) +
  scale_fill_manual(values = c("0" = "#9ECAE1", "I" = "#FDD0A2","II" = "#F89866", 
                               "III" = "#FEF3BD", "IV" = "#FEFFED"))

# Mostrar el gráfico
print(grafico_violines_TNMR)






##### AQP8 vs. GradoHist -----------------------------------------------------
# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `GradoHist`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$GradoHist), ]


# Crear el gráfico
grafico_violines_gradohist <- ggplot(data_filtrada, aes(x = GradoHist, y = `ΔΔCT AQP8`, fill = GradoHist)) +
  geom_violin(trim = FALSE, color = "black", linetype = "solid", size = 0.6, alpha = 0.7) +
  geom_boxplot(
    width = 0.1, 
    outlier.shape = 21, 
    outlier.size = 2, 
    outlier.stroke = 0.5, 
    outlier.fill = "white", 
    color = "black",
    fill = "white"
  ) +
  labs(
    x = "Histologic Grade",
    y = expression(AQP8~({Delta*Delta*C[t]}))
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    axis.title.y = element_text(face = "bold", size = 14),
    axis.text = element_text(size = 12, color = "black"),
    panel.grid = element_blank(),
    axis.line = element_line(color = "black"),
    axis.ticks = element_line(color = "black"),
    legend.position = "none"
  ) +
  scale_fill_manual(values = c("Well differentiated"     = "#9ECAE1", 
                                "Moderately differentiated"= "#FDD0A2", 
                                "Poorly differentiated"   = "#F89866"))

# Mostrar el gráfico
print(grafico_violines_gradohist)



##### AQP8 vs. Histol_dif -----------------------------------------------------

# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `Histol_dif`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Histol_dif), ]

shapiro.test(data$`ΔΔCT AQP8`[data$Histol_dif == "Well differentiated"])
shapiro.test(data$`ΔΔCT AQP8`[data$Histol_dif == "Some grade of undifferentiation"])


wilcox.test(`ΔΔCT AQP8` ~ Histol_dif, data = data)

# Crear el grafico
grafico_violines_histol_dif <- ggplot(data_filtrada, aes(x = Histol_dif, y = `ΔΔCT AQP8`, fill = Histol_dif)) +
  geom_violin(trim = FALSE, color = "black", linetype = "solid", size = 0.6, alpha = 0.7) +
  geom_boxplot(
    width = 0.1,  
    outlier.shape = 21,  
    outlier.size = 2,  
    outlier.stroke = 0.5,  
    outlier.fill = "white",  
    color = "black",
    fill = "white"
  ) +
  labs(
    x = "Histologic Grade Grade",
    y = expression(AQP8~(2^{-Delta*Delta*C[t]}))
  ) +
  # Línea añadida para cambiar las etiquetas del eje X
  scale_x_discrete(labels = c("Well differentiated" = "WD", 
                              "Some grade of undifferentiation" = "SGU")) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    axis.title.y = element_text(face = "bold", size = 14),
    axis.text = element_text(size = 12, color = "black"),
    panel.grid = element_blank(),
    axis.line = element_line(color = "black"),
    axis.ticks = element_line(color = "black"),
    legend.position = "none"
  ) +
  scale_fill_manual(values = c("Well differentiated"           = "#9ECAE1",  
                               "Some grade of undifferentiation" = "#FDD0A2"))

# Para visualizar el gráfico
print(grafico_violines_histol_dif)




###############################################################################
# 3.1. Expression plots Western Blot
###############################################################################

# 2. Prepare factors and combined group (6 ordered groups)
data_wb <- data_wb %>%
  mutate(
    ESTADIO = factor(ESTADIO,
                     levels = c("Control", "Early", "Advanced")),
    
    # Now RIGHT colon first, then LEFT colon
    LOCALIZACIÓN = factor(LOCALIZACIÓN,
                          levels = c("DERECHO", "IZQUIERDO")),
    
    # Combined group: LOCATION_first + STAGE
    GRUPO_RAW = interaction(LOCALIZACIÓN, ESTADIO, sep = "_"),
    
    GRUPO = factor(
      GRUPO_RAW,
      levels = c(
        # RIGHT colon first
        "DERECHO_Control",
        "DERECHO_Early",
        "DERECHO_Advanced",
        # LEFT colon second
        "IZQUIERDO_Control",
        "IZQUIERDO_Early",
        "IZQUIERDO_Advanced"
      ),
      labels = c(
        "Control\nRight colon",
        "Early\nRight colon",
        "Advanced\nRight colon",
        "Control\nLeft colon",
        "Early\nLeft colon",
        "Advanced\nLeft colon"
      )
    )
  )

# 3. Normality check for each of the 6 groups
normality_test <- data_wb %>%
  group_by(GRUPO) %>%
  summarise(
    n         = n(),
    shapiro_p = shapiro.test(`EXP RELATIVA AQP8`)$p.value,
    normal    = shapiro_p > 0.05
  )

print(normality_test)

# 4. Global test + pairwise comparisons
if (all(normality_test$normal)) {
  
  global_result <- aov(`EXP RELATIVA AQP8` ~ GRUPO, data = data_wb) %>%
    summary()
  print("ANOVA results:")
  print(global_result)
  
  comparisons <- data_wb %>%
    tukey_hsd(`EXP RELATIVA AQP8` ~ GRUPO) %>%
    select(group1, group2, p.adj, p.adj.signif)
  
  test_used <- "ANOVA with post-hoc Tukey HSD"
  
} else {
  
  global_result <- kruskal.test(`EXP RELATIVA AQP8` ~ GRUPO, data = data_wb)
  print("Kruskal-Wallis results:")
  print(global_result)
  
  comparisons <- data_wb %>%
    dunn_test(`EXP RELATIVA AQP8` ~ GRUPO, 
              p.adjust.method = "none") %>%
    select(group1, group2, p.adj, p.adj.signif)
  
  test_used <- "Kruskal-Wallis with Dunn post-hoc + Bonferroni"
  
}

print(comparisons)
print(test_used)

# 5. Violin + jitter + boxplot (6 ordered violins)
pastel_colors_6 <- c("#C6DBEF", "#9ECAE1", "#FDCDAC",
                     "#FDD0A2", "#c7e9c0", "#ECFEE8")

p <- ggplot(data_wb,
            aes(x = GRUPO,
                y = `EXP RELATIVA AQP8`,
                fill = GRUPO)) +
  geom_violin(trim = FALSE,
              color = "black",
              size = 0.8,
              alpha = 0.5,
              width = 1) +
  geom_jitter(width = 0.12,
              alpha = 0.65,
              size = 1.8,
              shape = 21,
              color = "black",
              stroke = 0.6) +
  geom_boxplot(width = 0.1,
               outlier.shape = NA,
               color = "black",
               fill = "white",
               size = 0.6) +
  scale_y_log10() +
  labs(
    x = "",
    y = "Relative AQP8 expression"
  ) +
  scale_fill_manual(values = pastel_colors_6) +
  theme_minimal(base_size = 12) +
  theme(
    axis.title.y = element_text(face = "bold",
                                size = 14,
                                margin = margin(r = 15)),
    axis.text = element_text(size = 11, color = "black"),
    axis.text.x = element_text(angle = 45, hjust = 1),
    panel.grid = element_blank(),
    axis.line.x.bottom = element_line(color = "black", size = 0.8),
    axis.line.y.left   = element_line(color = "black", size = 0.8),
    axis.ticks.x.bottom = element_line(color = "black", size = 0.8),
    axis.ticks.y.left   = element_line(color = "black", size = 0.8),
    legend.position = "none",
    plot.margin = margin(20, 20, 20, 20)
  )

p




###############################################################################
# 4. Post-hoc analyses
###############################################################################
################################################################################
#
# SCRIPT COMPLETO: ANÁLISIS POST-HOC TNMR Y EXPORTACIÓN A WORD (VERSIÓN FINAL)
#
# Este script realiza un análisis ANOVA con una prueba post-hoc de Tukey,
# calcula estadísticas descriptivas y tamaños del efecto, y finalmente
# exporta una tabla con formato de publicación (estilo The Lancet, en inglés)
# a un archivo .docx.
#
################################################################################


################################################################################
# PASO 1: INSTALAR Y CARGAR LIBRERÍAS
# Ejecuta esta sección una vez para instalar los paquetes si no los tienes.
################################################################################
library(dplyr)
library(effsize)
library(flextable)
library(officer)


################################################################################
# PASO 2: CARGAR LOS DATOS
# He creado datos de ejemplo para que el script sea completamente funcional.
# **SUSTITUYE ESTA SECCIÓN CON TUS PROPIOS DATOS**, por ejemplo:
# data <- read.csv("mi_archivo_de_datos.csv")
################################################################################

################################################################################
# PASO 3: ANÁLISIS ESTADÍSTICO Y CÁLCULO DE MÉTRICAS (VERSIÓN CORREGIDA)
# Esta versión calcula la d de Cohen de forma más robusta para evitar valores faltantes.
################################################################################

# 3a. Filtrar datos para eliminar NAs (si los hubiera)
data_filtrada <- data %>%
  filter(!is.na(`ΔΔCT AQP8`) & !is.na(TNMR))

# 3b. Realizar ANOVA y prueba Tukey HSD
anova_model <- aov(`ΔΔCT AQP8` ~ TNMR, data = data_filtrada)
tukey_result <- TukeyHSD(anova_model, "TNMR")
tukeydf <- as.data.frame(tukey_result[["TNMR"]])
tukeydf <- tibble::rownames_to_column(tukeydf, "comparison")

# 3c. Calcular estadísticas descriptivas (N, Media, DE) para cada grupo
group_stats <- data_filtrada %>%
  group_by(TNMR) %>%
  summarise(
    n = n(),
    mean = mean(`ΔΔCT AQP8`),
    sd = sd(`ΔΔCT AQP8`),
    .groups = 'drop'
  )

# --- INICIO DE LA CORRECCIÓN ---
# 3d. Calcular d de Cohen iterando directamente sobre la tabla de Tukey para garantizar concordancia.
cohen_d_values <- sapply(tukeydf$comparison, function(comp) {
  # Extraer los dos grupos del nombre de la comparación (ej. "I-0")
  group_labels <- strsplit(comp, "-")[[1]]
  group2_label <- group_labels[1]
  group1_label <- group_labels[2]
  
  # Obtener los datos para cada grupo
  group1_data <- data_filtrada$`ΔΔCT AQP8`[data_filtrada$TNMR == group1_label]
  group2_data <- data_filtrada$`ΔΔCT AQP8`[data_filtrada$TNMR == group2_label]
  
  # Calcular d de Cohen. El orden es cohen.d(grupo_tratamiento, grupo_control)
  # Tukey calcula la diferencia como group2 - group1, así que usamos el mismo orden.
  if (length(group1_data) > 1 && length(group2_data) > 1) {
    d_value <- cohen.d(group2_data, group1_data, na.rm = TRUE)$estimate
    return(d_value)
  } else {
    return(NA) # Retorna NA si un grupo tiene n<2, lo cual previene errores.
  }
})

# Añadir los valores de d de Cohen calculados a la tabla de Tukey
tukeydf$cohen_d <- cohen_d_values
# --- FIN DE LA CORRECCIÓN ---


# 3e. Ensamblar la tabla de resultados final
final_publication_table <- left_join(tukeydf, group_stats, by = c("comparison" = "TNMR")) # This join is now simpler
final_publication_table <- tukeydf %>% # Start from the enhanced tukeydf
  mutate(
    group2 = sub("-.*", "", comparison),
    group1 = sub(".*-", "", comparison)
  ) %>%
  left_join(group_stats, by = c("group1" = "TNMR")) %>%
  rename(n_group1 = n, mean_group1 = mean, sd_group1 = sd) %>%
  left_join(group_stats, by = c("group2" = "TNMR")) %>%
  rename(n_group2 = n, mean_group2 = mean, sd_group2 = sd) %>%
  select(comparison, n_group1, mean_group1, sd_group1, n_group2, mean_group2, sd_group2, diff, lwr, upr, cohen_d, `p adj`) %>%
  rename(mean_difference = diff, ci_lower_95 = lwr, ci_upper_95 = upr, p_adjusted = `p adj`)


################################################################################
# PASO 4: CREAR Y GUARDAR LA TABLA EN UN ARCHIVO .DOCX
# Esta sección toma la tabla final y le da el formato para publicación.
################################################################################

# 4a. Preparar los datos para una presentación limpia en la tabla
table_for_docx <- final_publication_table %>%
  mutate(
    mean_diff_ci = sprintf("%.2f (%.2f to %.2f)", mean_difference, ci_lower_95, ci_upper_95),
    n_groups = paste0(n_group1, " / ", n_group2)
  ) %>%
  select(comparison, n_groups, mean_group1, sd_group1, mean_group2, sd_group2, mean_diff_ci, cohen_d, p_adjusted)

# 4b. Crear y formatear el objeto 'flextable' en inglés
ft <- flextable(table_for_docx) %>%
  set_header_labels(
    comparison = "Group Comparison",
    n_groups = "n (Group 1 / Group 2)",
    mean_group1 = "Mean ± SD (Group 1)",
    sd_group1 = "",
    mean_group2 = "Mean ± SD (Group 2)",
    sd_group2 = "",
    mean_diff_ci = "Mean Difference (95% CI)",
    cohen_d = "Cohen's d",
    p_adjusted = "Adjusted p-value"
  ) %>%
  compose(i = ~ !is.na(mean_group1), j = "mean_group1", value = as_paragraph(sprintf("%.2f ± %.2f", mean_group1, sd_group1))) %>%
  compose(i = ~ !is.na(mean_group2), j = "mean_group2", value = as_paragraph(sprintf("%.2f ± %.2f", mean_group2, sd_group2))) %>%
  theme_booktabs() %>%
  bold(part = "header", bold = TRUE) %>%
  align(align = "center", part = "all") %>%
  align(j = "comparison", align = "left", part = "all") %>%
  autofit() %>%
  width(j = "comparison", width = 1.8) %>%
  width(j = c("mean_group1", "mean_group2"), width = 1.2)

# 4c. Definir título y notas al pie en inglés
table_title <- "Table 1: Comparison of relative AQP8 gene expression (ΔΔCT) across clinical groups."
footnote1 <- "SD, Standard Deviation; CI, Confidence Interval."
footnote2 <- "p-values were adjusted using Tukey's honestly significant difference (HSD) test following a one-way ANOVA. A higher ΔΔCT value indicates higher relative gene expression."

# 4d. Crear el documento de Word y añadir los elementos (CON LA CORRECCIÓN)
doc <- read_docx() %>%
  body_add_par(value = table_title, style = "heading 1") %>%
  body_add_par("") %>%
  body_add_flextable(value = ft) %>%
  body_add_par("") %>%
  # --- INICIO DE LA CORRECCIÓN ---
  # Usamos el estilo "Normal" porque "footnote" no existe en la plantilla por defecto.
  body_add_par(value = footnote1, style = "Normal") %>%
  body_add_par(value = footnote2, style = "Normal")
# --- FIN DE LA CORRECCIÓN ---

# 4e. Guardar el documento final
file_path <- "Table1_TNMR_Results_Final.docx"
print(doc, target = file_path)

# Mensaje de confirmación
print(paste("Success! The final table has been saved to:", file_path))






# Post hoc for GradoHist using Dunn  -----------------------------------------

################################################################################
#
# SCRIPT COMPLETO: ANÁLISIS NO PARAMÉTRICO Y EXPORTACIÓN A WORD (VERSIÓN FINAL)
#
# Este script realiza una prueba de Kruskal-Wallis con una prueba post-hoc
# de Dunn, calcula estadísticas descriptivas (Mediana, IQR) y tamaños del
# efecto (r), y exporta una tabla con formato de publicación en inglés
# a un archivo .docx.
#
################################################################################


################################################################################
# PASO 1: INSTALAR Y CARGAR LIBRERÍAS
# El paquete 'FSA' es nuevo y necesario para la prueba de Dunn.
################################################################################
if (!require(dplyr)) install.packages("dplyr")
if (!require(FSA)) install.packages("FSA") # Para dunnTest()
if (!require(flextable)) install.packages("flextable")
if (!require(officer)) install.packages("officer")

library(dplyr)
library(FSA)
library(flextable)
library(officer)


################################################################################
# PASO 2: CARGAR LOS DATOS
# **SUSTITUYE ESTA SECCIÓN CON TUS PROPIOS DATOS**
# Nota: La nueva variable de agrupación es 'GradoHist'.
################################################################################

################################################################################
# PASO 3: ANÁLISIS ESTADÍSTICO Y CÁLCULO DE MÉTRICAS (NO PARAMÉTRICO)
################################################################################

# 3a. Filtrar datos para eliminar NAs de las nuevas variables
data_filtrada <- data %>%
  filter(!is.na(`ΔΔCT AQP8`) & !is.na(GradoHist))

# 3b. Realizar prueba de Kruskal-Wallis y prueba post-hoc de Dunn
kruskal_test <- kruskal.test(`ΔΔCT AQP8` ~ GradoHist, data = data_filtrada)
print(kruskal_test) # Imprimir el resultado general de Kruskal-Wallis

# --- INICIO DE LA CORRECCIÓN ---
# Se cambia la llamada a dunnTest para usar la sintaxis x= y g= que es más robusta.
dunn_result <- dunnTest(x = data_filtrada$`ΔΔCT AQP8`, g = data_filtrada$GradoHist, method = "bh")
# --- FIN DE LA CORRECCIÓN ---

dunndf <- as.data.frame(dunn_result$res)
# Renombrar columnas para claridad
dunndf <- dunndf %>%
  rename(
    comparison = Comparison,
    z_statistic = Z,
    p_unadjusted = P.unadj,
    p_adjusted = P.adj
  )

# 3c. Calcular estadísticas descriptivas NO PARAMÉTRICAS (N, Mediana, IQR)
group_stats <- data_filtrada %>%
  group_by(GradoHist) %>%
  summarise(
    n = n(),
    median = median(`ΔΔCT AQP8`),
    iqr = IQR(`ΔΔCT AQP8`),
    .groups = 'drop'
  )

# 3d. Calcular tamaño del efecto (r) a partir del estadístico Z de Dunn
# r = Z / sqrt(N), donde N es el número total de observaciones en los dos grupos comparados.
dunndf <- dunndf %>%
  mutate(
    group2 = sub(" - .*", "", comparison),
    group1 = sub(".* - ", "", comparison)
  ) %>%
  left_join(group_stats %>% select(GradoHist, n), by = c("group1" = "GradoHist")) %>%
  rename(n1 = n) %>%
  left_join(group_stats %>% select(GradoHist, n), by = c("group2" = "GradoHist")) %>%
  rename(n2 = n) %>%
  mutate(
    effect_size_r = abs(z_statistic) / sqrt(n1 + n2)
  ) %>%
  select(-group1, -group2, -n1, -n2) # Limpiar columnas auxiliares

# 3e. Ensamblar la tabla de resultados final
final_publication_table <- dunndf %>%
  mutate(
    group2 = sub(" - .*", "", comparison),
    group1 = sub(".* - ", "", comparison)
  ) %>%
  left_join(group_stats, by = c("group1" = "GradoHist")) %>%
  rename(n_group1 = n, median_group1 = median, iqr_group1 = iqr) %>%
  left_join(group_stats, by = c("group2" = "GradoHist")) %>%
  rename(n_group2 = n, median_group2 = median, iqr_group2 = iqr)


################################################################################
# PASO 4: CREAR Y GUARDAR LA TABLA EN UN ARCHIVO .DOCX
# Adaptado para los resultados no paramétricos.
################################################################################

# 4a. Preparar los datos para una presentación limpia en la tabla
table_for_docx <- final_publication_table %>%
  mutate(
    n_groups = paste0(n_group1, " / ", n_group2),
    median1_iqr = sprintf("%.2f (%.2f)", median_group1, iqr_group1),
    median2_iqr = sprintf("%.2f (%.2f)", median_group2, iqr_group2)
  ) %>%
  select(comparison, n_groups, median1_iqr, median2_iqr, z_statistic, effect_size_r, p_adjusted)

# 4b. Crear y formatear el objeto 'flextable' en inglés
ft <- flextable(table_for_docx) %>%
  set_header_labels(
    comparison = "Group Comparison",
    n_groups = "n (Group 1 / Group 2)",
    median1_iqr = "Median (IQR) Group 1",
    median2_iqr = "Median (IQR) Group 2",
    z_statistic = "Z-statistic",
    effect_size_r = "Effect Size (r)",
    p_adjusted = "Adjusted p-value"
  ) %>%
  theme_booktabs() %>%
  bold(part = "header", bold = TRUE) %>%
  align(align = "center", part = "all") %>%
  align(j = "comparison", align = "left", part = "all") %>%
  autofit() %>%
  colformat_double(j = c("z_statistic", "effect_size_r", "p_adjusted"), digits = 3) %>%
  width(j = "comparison", width = 1.8)

# 4c. Definir título y notas al pie en inglés, reflejando los nuevos métodos
table_title <- "Table 2: Pairwise comparisons of relative AQP8 gene expression (ΔΔCT) by histological grade."
footnote1 <- "IQR, Interquartile Range."
footnote2 <- "p-values represent pairwise comparisons following a Kruskal-Wallis test, adjusted using Dunn's test with the Benjamini-Hochberg method. Effect size (r) was calculated as Z/sqrt(N)."

# 4d. Crear el documento de Word y añadir los elementos
doc <- read_docx() %>%
  body_add_par(value = table_title, style = "heading 1") %>%
  body_add_par("") %>%
  body_add_flextable(value = ft) %>%
  body_add_par("") %>%
  body_add_par(value = footnote1, style = "Normal") %>%
  body_add_par(value = footnote2, style = "Normal")

# 4e. Guardar el documento final
file_path <- "Table2_GradoHist_Results_Final.docx"
print(doc, target = file_path)

# Mensaje de confirmación
print(paste("Success! The non-parametric analysis table has been saved to:", file_path))










###############################################################################
# 5. ROC analysis
###############################################################################

######## AQP8 and Histol_dif ----------------------------------------------------

# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `Histol_dif`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Histol_dif), ]

# Install and load the necessary package (if not already installed):
# install.packages("pROC")
library(pROC)

# Ensure the factor levels are correctly ordered:
# "Well differentiated" = negative group, "Some grade of undifferentiation" = positive group
data_filtrada$Histol_dif <- relevel(data_filtrada$Histol_dif, ref = "Well differentiated")

# Build the ROC object
roc_obj <- roc(
  response = data_filtrada$Histol_dif,
  predictor = data_filtrada$`2^ΔΔCT AQP8`,
  levels = c("Well differentiated", "Some grade of undifferentiation"),
  direction = "<"
)

# Print the AUC
cat("AUC:", roc_obj$auc, "\n")

# Extract the optimal cutoff, sensitivity, and specificity
best_coords <- coords(roc_obj, x = "best", ret = c("threshold", "sensitivity", "specificity"))
cat("Optimal Threshold:", best_coords$threshold, "\n")
cat("Sensitivity:", best_coords$sensitivity, "\n")
cat("Specificity:", best_coords$specificity, "\n")

# Alternatively, if you must index, convert to a character or numeric vector:
cat("Optimal Threshold:", as.numeric(best_coords["threshold"]), "\n")

# Minimalistic ROC plot
plot.roc(
  roc_obj,
  col = "#1f77b4",
  lwd = 3,
  legacy.axes = TRUE,
  main = "",
  xlab = "1 - Specificity",
  ylab = "Sensitivity"
)




######## AQP8 and Progres --------------------------------------------------

# Filtrar datos para eliminar los NA de la variable `2^ΔΔCT AQP8` y `Progres`
data_filtrada <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Progres), ]

# Convertir la variable Progres en factor para asegurarnos del orden
data_filtrada$Progres <- factor(data_filtrada$Progres, levels = c("No", "Yes"))

# Build the ROC object
roc_obj_progres <- roc(
  response = data_filtrada$Progres,
  predictor = data_filtrada$`2^ΔΔCT AQP8`,
  levels = c("No", "Yes"),
  direction = "<"
)

# Print the AUC
cat("AUC for Progres:", roc_obj_progres$auc, "\n")
# Extract the optimal cutoff, sensitivity, and specificity
best_coords_progres <- coords(roc_obj_progres, x = "best", ret = c("threshold", "sensitivity", "specificity"))
cat("Optimal Threshold for Progres:", best_coords_progres$threshold, "\n")
cat("Sensitivity for Progres:", best_coords_progres$sensitivity, "\n")
cat("Specificity for Progres:", best_coords_progres$specificity, "\n")

# Minimalistic ROC plot for Progres
plot.roc(
  roc_obj_progres,
  col = "#1f77b4",
  lwd = 3,
  legacy.axes = TRUE,
  main = "",
  xlab = "1 - Specificity",
  ylab = "Sensitivity"
)







###############################################################################
# 6. Logistic Regression models
###############################################################################

###############################################################################
# Stratified Analysis of AQP8 and Histological Differentiation by Subgroup
###############################################################################

#-------------------------------------------------------------------------------
# 0. PREÁMBULO Y LIBRERÍAS
#-------------------------------------------------------------------------------
library(ggplot2)
library(dplyr)
library(tibble)
library(broom)
library(purrr)   # For map functions

#-------------------------------------------------------------------------------
# 1. PREPARACIÓN DE DATOS
#-------------------------------------------------------------------------------

# Filtrar datos y preparar factores para el análisis
data_histol <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Histol_dif) & !is.na(data$group), ]

# Asegurar que las variables son factores con niveles adecuados
data_histol$Histol_dif <- factor(data_histol$Histol_dif, 
                                 levels = c("Well differentiated", "Some grade of undifferentiation"))

# Convertir grupo a factor sin cambiar el orden
data_histol$group <- factor(data_histol$group)

# Ver la tabla de contingencia
print("Tabla de contingencia entre grupo y diferenciación histológica:")
print(table(data_histol$group, data_histol$Histol_dif))

#-------------------------------------------------------------------------------
# 2. MODELOS LOGÍSTICOS ESTRATIFICADOS POR GRUPO
#-------------------------------------------------------------------------------

# Obtener los subgrupos únicos
subgroups <- unique(data_histol$group)

# Función para ajustar los modelos para un subgrupo específico
fit_models_for_subgroup <- function(subgroup_name) {
  # Filtrar datos para este subgrupo
  subgroup_data <- data_histol %>% filter(group == subgroup_name)
  
  # Modelo 1: Sin ajustar (solo AQP8)
  model1 <- tryCatch({
    glm(Histol_dif ~ `ΔΔCT AQP8`, data = subgroup_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 1 para", subgroup_name, ":", e$message))
    return(NULL)
  })
  
  # Modelo 2: Ajustado por edad y sexo
  model2 <- tryCatch({
    glm(Histol_dif ~ `ΔΔCT AQP8` + Edad + Sexo, data = subgroup_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 2 para", subgroup_name, ":", e$message))
    return(NULL)
  })
  
  # Devolver los modelos en una lista
  list(
    subgroup = subgroup_name,
    model1 = model1,
    model2 = model2
  )
}

# Ajustar los modelos para cada subgrupo
all_models <- lapply(subgroups, fit_models_for_subgroup)
names(all_models) <- subgroups

#-------------------------------------------------------------------------------
# 3. EXTRAER ODDS RATIOS E INTERVALOS DE CONFIANZA
#-------------------------------------------------------------------------------

# Función para extraer OR para AQP8 de un modelo
extract_or_from_model <- function(model, subgroup_name, model_type) {
  if (is.null(model)) {
    # Si el modelo es NULL (falló), devolver NA
    return(data.frame(
      subgroup = subgroup_name,
      model = model_type,
      OR = NA,
      LCI = NA,
      UCI = NA,
      p_value = NA
    ))
  }
  
  # Extraer coeficientes y CI
  coefs <- coef(summary(model))
  aqp8_index <- which(rownames(coefs) == "`ΔΔCT AQP8`")
  
  if (length(aqp8_index) == 0) {
    warning(paste("No se encontró el coeficiente para AQP8 en", subgroup_name, model_type))
    return(NULL)
  }
  
  aqp8_coef <- coefs[aqp8_index, 1]
  aqp8_se <- coefs[aqp8_index, 2]
  aqp8_p <- coefs[aqp8_index, 4]
  
  # Calcular OR e intervalos de confianza
  or <- exp(aqp8_coef)
  lci <- exp(aqp8_coef - 1.96 * aqp8_se)
  uci <- exp(aqp8_coef + 1.96 * aqp8_se)
  
  # Devolver como data frame
  data.frame(
    subgroup = subgroup_name,
    model = model_type,
    OR = or,
    LCI = lci,
    UCI = uci,
    p_value = aqp8_p
  )
}

# Extraer OR para cada modelo y subgrupo
or_results <- data.frame()

for (i in 1:length(all_models)) {
  models <- all_models[[i]]
  subgroup <- models$subgroup
  
  # Extraer OR del modelo 1 (sin ajustar)
  or1 <- extract_or_from_model(models$model1, subgroup, "Model 1 (Unadjusted)")
  
  # Extraer OR del modelo 2 (ajustado)
  or2 <- extract_or_from_model(models$model2, subgroup, "Model 2 (Adjusted)")
  
  # Combinar resultados
  or_results <- rbind(or_results, or1, or2)
}

# Ordenar resultados para una mejor visualización
or_results <- or_results %>%
  arrange(subgroup, model)

# Mostrar la tabla de resultados
print("Odds Ratios for ΔΔCT AQP8 by Subgroup:")
print(or_results %>% 
        mutate(
          OR_formatted = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
          p_value_formatted = sprintf("%.4f", p_value)
        ) %>%
        select(subgroup, model, OR_formatted, p_value_formatted))

#-------------------------------------------------------------------------------
# 4. PREPARAR DATOS PARA LA VISUALIZACIÓN
#-------------------------------------------------------------------------------

# Crear factores ordenados para la visualización
plot_data <- or_results %>%
  mutate(
    subgroup = factor(subgroup, levels = c(
      "Distal colon - Early", 
      "Distal colon - Advanced",
      "Proximal colon - Early", 
      "Proximal colon - Advanced"
    )),
    model = factor(model, levels = c(
      "Model 1 (Unadjusted)", 
      "Model 2 (Adjusted)"
    ))
  )

#-------------------------------------------------------------------------------
# 5. CREAR EL FOREST PLOT
#-------------------------------------------------------------------------------

# Determinar límites para el gráfico
max_value <- max(plot_data$UCI, na.rm = TRUE)
max_limit <- min(max(10, ceiling(max_value * 1.1)), 100) # Limitar a 100 para evitar rangos extremos

forest_plot_aqp8 <- ggplot(plot_data, 
                           aes(x = OR, xmin = LCI, xmax = UCI, y = subgroup, color = model)) +
  
  # Línea de referencia en OR=1 (sin efecto)
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales
  geom_errorbarh(aes(height = 0.15), position = position_dodge(width = 0.5)) +
  
  # Puntos para los OR (usando cuadrados como en tu gráfico original)
  geom_point(size = 3, position = position_dodge(width = 0.5), shape = 15) +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50, 100),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  
  # Controlar el rango visible
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Odds Ratios for Histological Undifferentiation by AQP8 Expression",
    subtitle = "Stratified by Colon Cancer Location and Stage",
    y = NULL,
    color = "Model"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12)
  ) +
  
  # Escala de color personalizada
  scale_color_manual(values = c("Model 1 (Unadjusted)" = "#3A86FF", 
                                "Model 2 (Adjusted)" = "#FF006E"))

#-------------------------------------------------------------------------------
# 6. MOSTRAR EL GRÁFICO
#-------------------------------------------------------------------------------
print(forest_plot_aqp8)

#-------------------------------------------------------------------------------
# 7. CREAR TABLA DETALLADA DE RESULTADOS
#-------------------------------------------------------------------------------

# Formatear los resultados para una tabla más legible
formatted_results <- or_results %>%
  mutate(
    OR_CI = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
    p_value_formatted = ifelse(p_value < 0.001, 
                               "<0.001", 
                               sprintf("%.3f", p_value))
  ) %>%
  select(Subgroup = subgroup, Model = model, `OR (95% CI)` = OR_CI, `p-value` = p_value_formatted) %>%
  arrange(Subgroup, Model)

# Imprimir la tabla formateada
print("Tabla detallada de resultados:")
print(formatted_results, row.names = FALSE)

#-------------------------------------------------------------------------------
# 8. ANÁLISIS ADICIONAL: COMPARACIÓN DE EFECTOS ENTRE SUBGRUPOS
#-------------------------------------------------------------------------------

# Ajustar un modelo con interacción entre AQP8 y grupo
# Esto prueba si el efecto de AQP8 varía significativamente entre los subgrupos
interaction_model <- glm(Histol_dif ~ `ΔΔCT AQP8` * group, 
                         data = data_histol, 
                         family = binomial)

cat("\nTest de interacción: ¿El efecto de AQP8 varía entre subgrupos?\n")
print(anova(interaction_model, test = "LRT"))

# Comparar con modelo sin interacción
main_effects_model <- glm(Histol_dif ~ `ΔΔCT AQP8` + group, 
                          data = data_histol, 
                          family = binomial)

cat("\nComparación de modelos con y sin interacción:\n")
print(anova(main_effects_model, interaction_model, test = "Chisq"))

# Imprimir resumen del modelo de interacción
cat("\nResumen del modelo de interacción:\n")
print(summary(interaction_model))



















###############################################################################
# Stratified Analysis of AQP8 and Histological Differentiation by Anatomical Subtype
###############################################################################

#-------------------------------------------------------------------------------
# 0. PREÁMBULO Y LIBRERÍAS
#-------------------------------------------------------------------------------
library(ggplot2)
library(dplyr)
library(tibble)
library(broom)
library(purrr)

#-------------------------------------------------------------------------------
# 1. PREPARACIÓN DE DATOS
#-------------------------------------------------------------------------------

# Filtrar datos y preparar factores para el análisis
data_histol <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Histol_dif) & !is.na(data$Subtipoanatóm), ]

# Asegurar que las variables son factores con niveles adecuados
data_histol$Histol_dif <- factor(data_histol$Histol_dif, 
                                 levels = c("Well differentiated", "Some grade of undifferentiation"))

# Convertir subtipo anatómico a factor sin cambiar el orden
data_histol$Subtipoanatóm <- factor(data_histol$Subtipoanatóm)

# Ver la tabla de contingencia
print("Tabla de contingencia entre subtipo anatómico y diferenciación histológica:")
print(table(data_histol$Subtipoanatóm, data_histol$Histol_dif))

#-------------------------------------------------------------------------------
# 2. MODELOS LOGÍSTICOS ESTRATIFICADOS POR SUBTIPO ANATÓMICO
#-------------------------------------------------------------------------------

# Obtener los subtipos anatómicos únicos
subtypes <- unique(data_histol$Subtipoanatóm)

# Función para ajustar los modelos para un subtipo anatómico específico
fit_models_for_subtype <- function(subtype_name) {
  # Filtrar datos para este subtipo
  subtype_data <- data_histol %>% filter(Subtipoanatóm == subtype_name)
  
  # Modelo 1: Sin ajustar (solo AQP8)
  model1 <- tryCatch({
    glm(Histol_dif ~ `ΔΔCT AQP8`, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 1 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Modelo 2: Ajustado por edad y sexo
  model2 <- tryCatch({
    glm(Histol_dif ~ `ΔΔCT AQP8` + Edad + Sexo, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 2 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Devolver los modelos en una lista
  list(
    subtype = subtype_name,
    model1 = model1,
    model2 = model2
  )
}

# Ajustar los modelos para cada subtipo anatómico
all_models <- lapply(subtypes, fit_models_for_subtype)
names(all_models) <- subtypes

#-------------------------------------------------------------------------------
# 3. EXTRAER ODDS RATIOS E INTERVALOS DE CONFIANZA
#-------------------------------------------------------------------------------

# Función para extraer OR para AQP8 de un modelo
extract_or_from_model <- function(model, subtype_name, model_type) {
  if (is.null(model)) {
    # Si el modelo es NULL (falló), devolver NA
    return(data.frame(
      subtype = subtype_name,
      model = model_type,
      OR = NA,
      LCI = NA,
      UCI = NA,
      p_value = NA
    ))
  }
  
  # Extraer coeficientes y CI
  coefs <- coef(summary(model))
  aqp8_index <- which(rownames(coefs) == "`ΔΔCT AQP8`")
  
  if (length(aqp8_index) == 0) {
    warning(paste("No se encontró el coeficiente para AQP8 en", subtype_name, model_type))
    return(NULL)
  }
  
  aqp8_coef <- coefs[aqp8_index, 1]
  aqp8_se <- coefs[aqp8_index, 2]
  aqp8_p <- coefs[aqp8_index, 4]
  
  # Calcular OR e intervalos de confianza
  or <- exp(aqp8_coef)
  lci <- exp(aqp8_coef - 1.96 * aqp8_se)
  uci <- exp(aqp8_coef + 1.96 * aqp8_se)
  
  # Devolver como data frame
  data.frame(
    subtype = subtype_name,
    model = model_type,
    OR = or,
    LCI = lci,
    UCI = uci,
    p_value = aqp8_p
  )
}

# Extraer OR para cada modelo y subtipo
or_results <- data.frame()

for (i in 1:length(all_models)) {
  models <- all_models[[i]]
  subtype <- models$subtype
  
  # Extraer OR del modelo 1 (sin ajustar)
  or1 <- extract_or_from_model(models$model1, subtype, "Model 1 (Unadjusted)")
  
  # Extraer OR del modelo 2 (ajustado)
  or2 <- extract_or_from_model(models$model2, subtype, "Model 2 (Adjusted)")
  
  # Combinar resultados
  or_results <- rbind(or_results, or1, or2)
}

# Ordenar resultados para una mejor visualización
or_results <- or_results %>%
  arrange(subtype, model)

# Mostrar la tabla de resultados
print("Odds Ratios for ΔΔCT AQP8 by Anatomical Subtype:")
print(or_results %>% 
        mutate(
          OR_formatted = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
          p_value_formatted = sprintf("%.4f", p_value)
        ) %>%
        select(subtype, model, OR_formatted, p_value_formatted))

#-------------------------------------------------------------------------------
# 4. ANÁLISIS DE MODELO GLOBAL CON INTERACCIÓN
#-------------------------------------------------------------------------------

# Ajustar modelo con interacción entre AQP8 y subtipo anatómico
interaction_model <- glm(Histol_dif ~ `ΔΔCT AQP8` * Subtipoanatóm, 
                         data = data_histol, 
                         family = binomial)

# Modelo sin interacción para comparación
main_effects_model <- glm(Histol_dif ~ `ΔΔCT AQP8` + Subtipoanatóm, 
                          data = data_histol, 
                          family = binomial)

cat("\nComparación de modelos con y sin interacción:\n")
print(anova(main_effects_model, interaction_model, test = "Chisq"))

cat("\nResumen del modelo de interacción:\n")
print(summary(interaction_model))

#-------------------------------------------------------------------------------
# 5. PREPARAR DATOS PARA LA VISUALIZACIÓN
#-------------------------------------------------------------------------------

# Crear factores ordenados para la visualización
plot_data <- or_results %>%
  mutate(
    subtype = factor(subtype, levels = unique(subtypes)),
    model = factor(model, levels = c("Model 1 (Unadjusted)", "Model 2 (Adjusted)"))
  )

#-------------------------------------------------------------------------------
# 6. CREAR EL FOREST PLOT
#-------------------------------------------------------------------------------

# Determinar límites para el gráfico
max_value <- max(plot_data$UCI, na.rm = TRUE)
max_limit <- min(max(10, ceiling(max_value * 1.1)), 100)

# Crear un forest plot más legible para solo 2 grupos
forest_plot_aqp8 <- ggplot(plot_data, 
                           aes(x = OR, xmin = LCI, xmax = UCI, y = model, color = subtype)) +
  
  # Línea de referencia en OR=1
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales y puntos
  geom_errorbarh(aes(height = 0.2), position = position_dodge(width = 0.5)) +
  geom_point(size = 4, position = position_dodge(width = 0.5), shape = 15) +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Effect of AQP8 Expression on Histological Undifferentiation",
    subtitle = "Stratified by Anatomical Location",
    y = NULL,
    color = "Anatomical Subtype"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12)
  ) +
  
  # Escala de color personalizada - colores distintivos para los dos subtipos
  scale_color_manual(values = c("Distal colon" = "#3A86FF", 
                                "Proximal colon" = "#FF006E"))

#-------------------------------------------------------------------------------
# 7. MOSTRAR EL GRÁFICO
#-------------------------------------------------------------------------------
print(forest_plot_aqp8)

#-------------------------------------------------------------------------------
# 8. TABLA DETALLADA DE RESULTADOS
#-------------------------------------------------------------------------------

# Formatear resultados para una tabla más completa
formatted_results <- or_results %>%
  mutate(
    OR_CI = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
    p_value_formatted = ifelse(p_value < 0.001, "<0.001", sprintf("%.3f", p_value)),
    n_samples = sapply(subtype, function(s) sum(!is.na(data_histol$Histol_dif[data_histol$Subtipoanatóm == s])))
  ) %>%
  select(
    `Anatomical Subtype` = subtype, 
    Model = model, 
    `Sample Size` = n_samples,
    `OR (95% CI)` = OR_CI, 
    `p-value` = p_value_formatted
  ) %>%
  arrange(`Anatomical Subtype`, Model)

# Imprimir la tabla formateada
print("Detailed Results Table:")
print(formatted_results, row.names = FALSE)

#-------------------------------------------------------------------------------
# 9. MODELO COMBINADO PARA COMPARACIÓN DIRECTA
#-------------------------------------------------------------------------------

# Ajustar un modelo combinado para probar formalmente si hay diferencia entre los subtipos
# Este modelo ajusta por el efecto principal del subtipo anatómico y la interacción
combined_model <- glm(Histol_dif ~ `ΔΔCT AQP8` + Subtipoanatóm + `ΔΔCT AQP8`:Subtipoanatóm + Edad + Sexo, 
                      data = data_histol, 
                      family = binomial)

cat("\nModelo combinado con interacción (ajustado por edad y sexo):\n")
print(summary(combined_model))

# Extraer el término de interacción para una interpretación directa
interaction_term <- coef(summary(combined_model))[grep(":", rownames(coef(summary(combined_model)))), ]

cat("\nTérmino de interacción (diferencia en el efecto de AQP8 entre subtipos):\n")
print(interaction_term)
cat("\nOdds Ratio para la interacción:", exp(interaction_term[1]), "\n")













###############################################################################
# Stratified Analysis of AQP8 and Metastasis Risk by Anatomical Subtype
###############################################################################

#-------------------------------------------------------------------------------
# 0. PREÁMBULO Y LIBRERÍAS
#-------------------------------------------------------------------------------
library(ggplot2)
library(dplyr)
library(tibble)
library(broom)
library(purrr)

#-------------------------------------------------------------------------------
# 1. PREPARACIÓN DE DATOS
#-------------------------------------------------------------------------------

# Filtrar datos y preparar factores para el análisis
data_meta <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$metastasis) & !is.na(data$Subtipoanatóm), ]

# Asegurar que las variables son factores con niveles adecuados
data_meta$metastasis <- factor(data_meta$metastasis, levels = c("No", "Yes"))

# Convertir subtipo anatómico a factor
data_meta$Subtipoanatóm <- factor(data_meta$Subtipoanatóm)

# Ver la tabla de contingencia
print("Tabla de contingencia entre subtipo anatómico y metástasis:")
print(table(data_meta$Subtipoanatóm, data_meta$metastasis))

#-------------------------------------------------------------------------------
# 2. MODELOS LOGÍSTICOS ESTRATIFICADOS POR SUBTIPO ANATÓMICO
#-------------------------------------------------------------------------------

# Obtener los subtipos anatómicos únicos
subtypes <- unique(data_meta$Subtipoanatóm)

# Función para ajustar los modelos para un subtipo anatómico específico
fit_models_for_subtype <- function(subtype_name) {
  # Filtrar datos para este subtipo
  subtype_data <- data_meta %>% filter(Subtipoanatóm == subtype_name)
  
  # Modelo 1: Sin ajustar (solo AQP8)
  model1 <- tryCatch({
    glm(metastasis ~ `ΔΔCT AQP8`, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 1 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Modelo 2: Ajustado por edad y sexo
  model2 <- tryCatch({
    glm(metastasis ~ `ΔΔCT AQP8` + Edad + Sexo, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 2 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Devolver los modelos en una lista
  list(
    subtype = subtype_name,
    model1 = model1,
    model2 = model2
  )
}

# Ajustar los modelos para cada subtipo anatómico
all_models <- lapply(subtypes, fit_models_for_subtype)
names(all_models) <- subtypes

#-------------------------------------------------------------------------------
# 3. EXTRAER ODDS RATIOS E INTERVALOS DE CONFIANZA
#-------------------------------------------------------------------------------

# Función para extraer OR para AQP8 de un modelo
extract_or_from_model <- function(model, subtype_name, model_type) {
  if (is.null(model)) {
    # Si el modelo es NULL (falló), devolver NA
    return(data.frame(
      subtype = subtype_name,
      model = model_type,
      OR = NA,
      LCI = NA,
      UCI = NA,
      p_value = NA
    ))
  }
  
  # Extraer coeficientes y CI
  coefs <- coef(summary(model))
  aqp8_index <- which(rownames(coefs) == "`ΔΔCT AQP8`")
  
  if (length(aqp8_index) == 0) {
    warning(paste("No se encontró el coeficiente para AQP8 en", subtype_name, model_type))
    return(NULL)
  }
  
  aqp8_coef <- coefs[aqp8_index, 1]
  aqp8_se <- coefs[aqp8_index, 2]
  aqp8_p <- coefs[aqp8_index, 4]
  
  # Calcular OR e intervalos de confianza
  or <- exp(aqp8_coef)
  lci <- exp(aqp8_coef - 1.96 * aqp8_se)
  uci <- exp(aqp8_coef + 1.96 * aqp8_se)
  
  # Devolver como data frame
  data.frame(
    subtype = subtype_name,
    model = model_type,
    OR = or,
    LCI = lci,
    UCI = uci,
    p_value = aqp8_p
  )
}

# Extraer OR para cada modelo y subtipo
or_results <- data.frame()

for (i in 1:length(all_models)) {
  models <- all_models[[i]]
  subtype <- models$subtype
  
  # Extraer OR del modelo 1 (sin ajustar)
  or1 <- extract_or_from_model(models$model1, subtype, "Model 1 (Unadjusted)")
  
  # Extraer OR del modelo 2 (ajustado)
  or2 <- extract_or_from_model(models$model2, subtype, "Model 2 (Adjusted)")
  
  # Combinar resultados
  or_results <- rbind(or_results, or1, or2)
}

# Ordenar resultados para una mejor visualización
or_results <- or_results %>%
  arrange(subtype, model)

# Mostrar la tabla de resultados
print("Odds Ratios for ΔΔCT AQP8 by Anatomical Subtype (Metastasis Risk):")
print(or_results %>% 
        mutate(
          OR_formatted = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
          p_value_formatted = sprintf("%.4f", p_value)
        ) %>%
        select(subtype, model, OR_formatted, p_value_formatted))

#-------------------------------------------------------------------------------
# 4. ANÁLISIS DE MODELO GLOBAL CON INTERACCIÓN
#-------------------------------------------------------------------------------

# Ajustar modelo con interacción entre AQP8 y subtipo anatómico
interaction_model <- glm(metastasis ~ `ΔΔCT AQP8` * Subtipoanatóm, 
                         data = data_meta, 
                         family = binomial)

# Modelo sin interacción para comparación
main_effects_model <- glm(metastasis ~ `ΔΔCT AQP8` + Subtipoanatóm, 
                          data = data_meta, 
                          family = binomial)

cat("\nComparación de modelos con y sin interacción:\n")
print(anova(main_effects_model, interaction_model, test = "Chisq"))

cat("\nResumen del modelo de interacción:\n")
print(summary(interaction_model))

#-------------------------------------------------------------------------------
# 5. PREPARAR DATOS PARA LA VISUALIZACIÓN
#-------------------------------------------------------------------------------

# Crear factores ordenados para la visualización
plot_data <- or_results %>%
  mutate(
    subtype = factor(subtype, levels = unique(subtypes)),
    model = factor(model, levels = c("Model 1 (Unadjusted)", "Model 2 (Adjusted)"))
  )

#-------------------------------------------------------------------------------
# 6. CREAR EL FOREST PLOT
#-------------------------------------------------------------------------------

# Determinar límites para el gráfico
max_value <- max(plot_data$UCI, na.rm = TRUE)
max_limit <- min(max(10, ceiling(max_value * 1.1)), 100)

# Crear un forest plot más legible para solo 2 grupos
forest_plot_aqp8 <- ggplot(plot_data, 
                           aes(x = OR, xmin = LCI, xmax = UCI, y = model, color = subtype)) +
  
  # Línea de referencia en OR=1
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales y puntos
  geom_errorbarh(aes(height = 0.2), position = position_dodge(width = 0.5)) +
  geom_point(size = 4, position = position_dodge(width = 0.5), shape = 15) +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Effect of AQP8 Expression on Metastasis Risk",
    subtitle = "Stratified by Anatomical Location",
    y = NULL,
    color = "Anatomical Subtype"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12)
  ) +
  
  # Escala de color personalizada - colores distintivos para los dos subtipos
  scale_color_manual(values = c("Distal colon" = "#3A86FF", 
                                "Proximal colon" = "#FF006E"))

#-------------------------------------------------------------------------------
# 7. MOSTRAR EL GRÁFICO
#-------------------------------------------------------------------------------
print(forest_plot_aqp8)

#-------------------------------------------------------------------------------
# 8. TABLA DETALLADA DE RESULTADOS
#-------------------------------------------------------------------------------

# Calcular n y proporción de eventos para cada subtipo
subtype_summary <- data_meta %>%
  group_by(Subtipoanatóm) %>%
  summarise(
    total_n = n(),
    events = sum(metastasis == "Yes"),
    event_rate = sprintf("%.1f%%", 100 * sum(metastasis == "Yes") / n())
  )

# Formatear resultados para una tabla más completa
formatted_results <- or_results %>%
  mutate(
    OR_CI = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
    p_value_formatted = ifelse(p_value < 0.001, "<0.001", sprintf("%.3f", p_value))
  ) %>%
  left_join(subtype_summary, by = c("subtype" = "Subtipoanatóm")) %>%
  select(
    `Anatomical Subtype` = subtype, 
    Model = model, 
    `Sample Size` = total_n,
    `Events (%)` = event_rate,
    `OR (95% CI)` = OR_CI, 
    `p-value` = p_value_formatted
  ) %>%
  arrange(`Anatomical Subtype`, Model)

# Imprimir la tabla formateada
print("Detailed Results Table for Metastasis Risk:")
print(formatted_results, row.names = FALSE)

#-------------------------------------------------------------------------------
# 9. MODELO COMBINADO PARA COMPARACIÓN DIRECTA
#-------------------------------------------------------------------------------

# Ajustar un modelo combinado para probar formalmente si hay diferencia entre los subtipos
combined_model <- glm(metastasis ~ `ΔΔCT AQP8` + Subtipoanatóm + `ΔΔCT AQP8`:Subtipoanatóm + Edad + Sexo, 
                      data = data_meta, 
                      family = binomial)

cat("\nModelo combinado con interacción (ajustado por edad y sexo):\n")
print(summary(combined_model))

# Extraer el término de interacción para una interpretación directa
interaction_term <- coef(summary(combined_model))[grep(":", rownames(coef(summary(combined_model)))), ]

cat("\nTérmino de interacción (diferencia en el efecto de AQP8 entre subtipos):\n")
print(interaction_term)
cat("\nOdds Ratio para la interacción:", exp(interaction_term[1]), "\n")

#-------------------------------------------------------------------------------
# 10. ANÁLISIS ADICIONAL: CURVAS ROC PARA EVALUAR CAPACIDAD PREDICTIVA
#-------------------------------------------------------------------------------

# Cargar paquete para análisis ROC
if (!requireNamespace("pROC", quietly = TRUE)) {
  # Si no está instalado, imprimir mensaje
  cat("\nEl paquete 'pROC' no está instalado. Para evaluar la capacidad predictiva con curvas ROC, instálalo con:\n")
  cat("install.packages('pROC')\n")
} else {
  library(pROC)
  
  cat("\nAnálisis de curvas ROC por subtipo anatómico:\n")
  
  # Función para crear y evaluar curvas ROC para un subtipo
  roc_analysis <- function(subtype_name) {
    subtype_data <- data_meta[data_meta$Subtipoanatóm == subtype_name, ]
    
    # Modelo 1: Sin ajustar
    model1 <- all_models[[subtype_name]]$model1
    if (!is.null(model1)) {
      pred1 <- predict(model1, type = "response")
      roc1 <- pROC::roc(subtype_data$metastasis, pred1, quiet = TRUE)
      auc1 <- pROC::auc(roc1)
      ci1 <- pROC::ci.auc(roc1)
      cat(sprintf("%s - Modelo 1 (sin ajustar): AUC = %.3f (%.3f-%.3f)\n", 
                  subtype_name, auc1, ci1[1], ci1[3]))
    }
    
    # Modelo 2: Ajustado
    model2 <- all_models[[subtype_name]]$model2
    if (!is.null(model2)) {
      pred2 <- predict(model2, type = "response")
      roc2 <- pROC::roc(subtype_data$metastasis, pred2, quiet = TRUE)
      auc2 <- pROC::auc(roc2)
      ci2 <- pROC::ci.auc(roc2)
      cat(sprintf("%s - Modelo 2 (ajustado): AUC = %.3f (%.3f-%.3f)\n", 
                  subtype_name, auc2, ci2[1], ci2[3]))
    }
  }
  
  # Evaluar cada subtipo
  for (subtype in subtypes) {
    roc_analysis(subtype)
  }
}














###############################################################################
# Stratified Analysis of AQP8 and Progression Risk by Anatomical Subtype
###############################################################################

#-------------------------------------------------------------------------------
# 0. PREÁMBULO Y LIBRERÍAS
#-------------------------------------------------------------------------------
library(ggplot2)
library(dplyr)
library(tibble)
library(broom)
library(purrr)

#-------------------------------------------------------------------------------
# 1. PREPARACIÓN DE DATOS
#-------------------------------------------------------------------------------

# Filtrar datos y preparar factores para el análisis
data_prog <- data[!is.na(data$`ΔΔCT AQP8`) & !is.na(data$Progres) & !is.na(data$Subtipoanatóm), ]

# Asegurar que las variables son factores con niveles adecuados
data_prog$Progres <- factor(data_prog$Progres, levels = c("No", "Yes"))

# Convertir subtipo anatómico a factor
data_prog$Subtipoanatóm <- factor(data_prog$Subtipoanatóm)

# Ver la tabla de contingencia
print("Tabla de contingencia entre subtipo anatómico y progresión tumoral:")
print(table(data_prog$Subtipoanatóm, data_prog$Progres))

#-------------------------------------------------------------------------------
# 2. MODELOS LOGÍSTICOS ESTRATIFICADOS POR SUBTIPO ANATÓMICO
#-------------------------------------------------------------------------------

# Obtener los subtipos anatómicos únicos
subtypes <- unique(data_prog$Subtipoanatóm)

# Función para ajustar los modelos para un subtipo anatómico específico
fit_models_for_subtype <- function(subtype_name) {
  # Filtrar datos para este subtipo
  subtype_data <- data_prog %>% filter(Subtipoanatóm == subtype_name)
  
  # Modelo 1: Sin ajustar (solo AQP8)
  model1 <- tryCatch({
    glm(Progres ~ `ΔΔCT AQP8`, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 1 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Modelo 2: Ajustado por edad y sexo
  model2 <- tryCatch({
    glm(Progres ~ `ΔΔCT AQP8` + Edad + Sexo, data = subtype_data, family = binomial)
  }, error = function(e) {
    message(paste("Error en modelo 2 para", subtype_name, ":", e$message))
    return(NULL)
  })
  
  # Devolver los modelos en una lista
  list(
    subtype = subtype_name,
    model1 = model1,
    model2 = model2
  )
}

# Ajustar los modelos para cada subtipo anatómico
all_models <- lapply(subtypes, fit_models_for_subtype)
names(all_models) <- subtypes

#-------------------------------------------------------------------------------
# 3. EXTRAER ODDS RATIOS E INTERVALOS DE CONFIANZA
#-------------------------------------------------------------------------------

# Función para extraer OR para AQP8 de un modelo
extract_or_from_model <- function(model, subtype_name, model_type) {
  if (is.null(model)) {
    # Si el modelo es NULL (falló), devolver NA
    return(data.frame(
      subtype = subtype_name,
      model = model_type,
      OR = NA,
      LCI = NA,
      UCI = NA,
      p_value = NA
    ))
  }
  
  # Extraer coeficientes y CI
  coefs <- coef(summary(model))
  aqp8_index <- which(rownames(coefs) == "`ΔΔCT AQP8`")
  
  if (length(aqp8_index) == 0) {
    warning(paste("No se encontró el coeficiente para AQP8 en", subtype_name, model_type))
    return(NULL)
  }
  
  aqp8_coef <- coefs[aqp8_index, 1]
  aqp8_se <- coefs[aqp8_index, 2]
  aqp8_p <- coefs[aqp8_index, 4]
  
  # Calcular OR e intervalos de confianza
  or <- exp(aqp8_coef)
  lci <- exp(aqp8_coef - 1.96 * aqp8_se)
  uci <- exp(aqp8_coef + 1.96 * aqp8_se)
  
  # Devolver como data frame
  data.frame(
    subtype = subtype_name,
    model = model_type,
    OR = or,
    LCI = lci,
    UCI = uci,
    p_value = aqp8_p
  )
}

# Extraer OR para cada modelo y subtipo
or_results <- data.frame()

for (i in 1:length(all_models)) {
  models <- all_models[[i]]
  subtype <- models$subtype
  
  # Extraer OR del modelo 1 (sin ajustar)
  or1 <- extract_or_from_model(models$model1, subtype, "Model 1 (Unadjusted)")
  
  # Extraer OR del modelo 2 (ajustado)
  or2 <- extract_or_from_model(models$model2, subtype, "Model 2 (Adjusted)")
  
  # Combinar resultados
  or_results <- rbind(or_results, or1, or2)
}

# Ordenar resultados para una mejor visualización
or_results <- or_results %>%
  arrange(subtype, model)

# Mostrar la tabla de resultados
print("Odds Ratios for ΔΔCT AQP8 by Anatomical Subtype (Progression Risk):")
print(or_results %>% 
        mutate(
          OR_formatted = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
          p_value_formatted = sprintf("%.4f", p_value)
        ) %>%
        select(subtype, model, OR_formatted, p_value_formatted))

#-------------------------------------------------------------------------------
# 4. ANÁLISIS DE MODELO GLOBAL CON INTERACCIÓN
#-------------------------------------------------------------------------------

# Ajustar modelo con interacción entre AQP8 y subtipo anatómico
interaction_model <- glm(Progres ~ `ΔΔCT AQP8` * Subtipoanatóm, 
                         data = data_prog, 
                         family = binomial)

# Modelo sin interacción para comparación
main_effects_model <- glm(Progres ~ `ΔΔCT AQP8` + Subtipoanatóm, 
                          data = data_prog, 
                          family = binomial)

cat("\nComparación de modelos con y sin interacción:\n")
print(anova(main_effects_model, interaction_model, test = "Chisq"))

cat("\nResumen del modelo de interacción:\n")
print(summary(interaction_model))

#-------------------------------------------------------------------------------
# 5. PREPARAR DATOS PARA LA VISUALIZACIÓN
#-------------------------------------------------------------------------------

# Crear factores ordenados para la visualización
plot_data <- or_results %>%
  mutate(
    subtype = factor(subtype, levels = unique(subtypes)),
    model = factor(model, levels = c("Model 1 (Unadjusted)", "Model 2 (Adjusted)"))
  )

#-------------------------------------------------------------------------------
# 6. CREAR EL FOREST PLOT
#-------------------------------------------------------------------------------

# Determinar límites para el gráfico
max_value <- max(plot_data$UCI, na.rm = TRUE)
max_limit <- min(max(10, ceiling(max_value * 1.1)), 100)

# Crear un forest plot más legible para solo 2 grupos
forest_plot_aqp8 <- ggplot(plot_data, 
                           aes(x = OR, xmin = LCI, xmax = UCI, y = model, color = subtype)) +
  
  # Línea de referencia en OR=1
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales y puntos
  geom_errorbarh(aes(height = 0.2), position = position_dodge(width = 0.5)) +
  geom_point(size = 4, position = position_dodge(width = 0.5), shape = 15) +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Effect of AQP8 Expression on Tumour Progression Risk",
    subtitle = "Stratified by Anatomical Location",
    y = NULL,
    color = "Anatomical Subtype"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12)
  ) +
  
  # Escala de color personalizada - colores distintivos para los dos subtipos
  scale_color_manual(values = c("Distal colon" = "#3A86FF", 
                                "Proximal colon" = "#FF006E"))

#-------------------------------------------------------------------------------
# 7. MOSTRAR EL GRÁFICO
#-------------------------------------------------------------------------------
print(forest_plot_aqp8)

#-------------------------------------------------------------------------------
# 8. TABLA DETALLADA DE RESULTADOS
#-------------------------------------------------------------------------------

# Calcular n y proporción de eventos para cada subtipo
subtype_summary <- data_prog %>%
  group_by(Subtipoanatóm) %>%
  summarise(
    total_n = n(),
    events = sum(Progres == "Yes"),
    event_rate = sprintf("%.1f%%", 100 * sum(Progres == "Yes") / n())
  )

# Formatear resultados para una tabla más completa
formatted_results <- or_results %>%
  mutate(
    OR_CI = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
    p_value_formatted = ifelse(p_value < 0.001, "<0.001", sprintf("%.3f", p_value))
  ) %>%
  left_join(subtype_summary, by = c("subtype" = "Subtipoanatóm")) %>%
  select(
    `Anatomical Subtype` = subtype, 
    Model = model, 
    `Sample Size` = total_n,
    `Progression Events (%)` = event_rate,
    `OR (95% CI)` = OR_CI, 
    `p-value` = p_value_formatted
  ) %>%
  arrange(`Anatomical Subtype`, Model)

# Imprimir la tabla formateada
print("Detailed Results Table for Tumour Progression Risk:")
print(formatted_results, row.names = FALSE)

#-------------------------------------------------------------------------------
# 9. MODELO COMBINADO PARA COMPARACIÓN DIRECTA
#-------------------------------------------------------------------------------

# Ajustar un modelo combinado para probar formalmente si hay diferencia entre los subtipos
combined_model <- glm(Progres ~ `ΔΔCT AQP8` + Subtipoanatóm + `ΔΔCT AQP8`:Subtipoanatóm + Edad + Sexo, 
                      data = data_prog, 
                      family = binomial)

cat("\nModelo combinado con interacción (ajustado por edad y sexo):\n")
print(summary(combined_model))

# Extraer el término de interacción para una interpretación directa
interaction_term <- coef(summary(combined_model))[grep(":", rownames(coef(summary(combined_model)))), ]

cat("\nTérmino de interacción (diferencia en el efecto de AQP8 entre subtipos):\n")
print(interaction_term)
cat("\nOdds Ratio para la interacción:", exp(interaction_term[1]), "\n")

#-------------------------------------------------------------------------------
# 10. ANÁLISIS ADICIONAL: CURVAS ROC PARA EVALUAR CAPACIDAD PREDICTIVA
#-------------------------------------------------------------------------------

# Cargar paquete para análisis ROC
if (!requireNamespace("pROC", quietly = TRUE)) {
  cat("\nEl paquete 'pROC' no está instalado. Para evaluar la capacidad predictiva con curvas ROC, instálalo con:\n")
  cat("install.packages('pROC')\n")
} else {
  library(pROC)
  
  cat("\nAnálisis de curvas ROC por subtipo anatómico:\n")
  
  # Función para crear y evaluar curvas ROC para un subtipo
  roc_analysis <- function(subtype_name) {
    subtype_data <- data_prog[data_prog$Subtipoanatóm == subtype_name, ]
    
    # Modelo 1: Sin ajustar
    model1 <- all_models[[subtype_name]]$model1
    if (!is.null(model1)) {
      pred1 <- predict(model1, type = "response")
      roc1 <- pROC::roc(subtype_data$Progres, pred1, quiet = TRUE)
      auc1 <- pROC::auc(roc1)
      ci1 <- pROC::ci.auc(roc1)
      cat(sprintf("%s - Modelo 1 (sin ajustar): AUC = %.3f (%.3f-%.3f)\n", 
                  subtype_name, auc1, ci1[1], ci1[3]))
    }
    
    # Modelo 2: Ajustado
    model2 <- all_models[[subtype_name]]$model2
    if (!is.null(model2)) {
      pred2 <- predict(model2, type = "response")
      roc2 <- pROC::roc(subtype_data$Progres, pred2, quiet = TRUE)
      auc2 <- pROC::auc(roc2)
      ci2 <- pROC::ci.auc(roc2)
      cat(sprintf("%s - Modelo 2 (ajustado): AUC = %.3f (%.3f-%.3f)\n", 
                  subtype_name, auc2, ci2[1], ci2[3]))
    }
  }
  
  # Evaluar cada subtipo
  for (subtype in subtypes) {
    roc_analysis(subtype)
  }
}




















###############################################################################
# Unified Analysis of AQP8 Expression on Multiple Colon Cancer Outcomes
###############################################################################

#-------------------------------------------------------------------------------
# 0. PREÁMBULO Y LIBRERÍAS
#-------------------------------------------------------------------------------
library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)
library(broom)
library(purrr)
library(forcats)
library(gridExtra)  # Para combinar tablas

#-------------------------------------------------------------------------------
# 1. FUNCIÓN PARA ANALIZAR CADA OUTCOME
#-------------------------------------------------------------------------------

analyze_outcome <- function(outcome_var, outcome_name, ref_level, event_level) {
  # Filtrar datos y preparar factores para el análisis
  outcome_formula <- as.formula(paste(outcome_var, "~ `ΔΔCT AQP8`"))
  outcome_adj_formula <- as.formula(paste(outcome_var, "~ `ΔΔCT AQP8` + Edad + Sexo"))
  
  # Datos filtrados para este outcome
  data_filtered <- data[!is.na(data[[outcome_var]]) & !is.na(data$`ΔΔCT AQP8`) & !is.na(data$Subtipoanatóm), ]
  
  # Asegurar factores con niveles correctos para cada variable
  data_filtered[[outcome_var]] <- factor(data_filtered[[outcome_var]], 
                                         levels = c(ref_level, event_level))
  data_filtered$Subtipoanatóm <- factor(data_filtered$Subtipoanatóm)
  
  # Mostrar tabla de contingencia
  cat(paste("\nTabla de contingencia para", outcome_name, ":\n"))
  print(table(data_filtered$Subtipoanatóm, data_filtered[[outcome_var]]))
  
  # Análisis por subtipo anatómico
  subtypes <- unique(data_filtered$Subtipoanatóm)
  
  or_results <- data.frame()
  
  for (subtype in subtypes) {
    # Filtrar por subtipo
    subtype_data <- data_filtered[data_filtered$Subtipoanatóm == subtype, ]
    
    # Verificar que hay suficientes datos para el modelo
    if (length(unique(subtype_data[[outcome_var]])) < 2) {
      message(paste("Advertencia: El subtipo", subtype, "para", outcome_name, 
                    "no tiene suficiente variabilidad (solo un nivel)"))
      next
    }
    
    # Modelo 1: Sin ajustar
    model1 <- tryCatch({
      glm(outcome_formula, data = subtype_data, family = binomial)
    }, error = function(e) {
      message(paste("Error en modelo 1 para", subtype, "-", outcome_name, ":", e$message))
      return(NULL)
    })
    
    # Modelo 2: Ajustado
    model2 <- tryCatch({
      glm(outcome_adj_formula, data = subtype_data, family = binomial)
    }, error = function(e) {
      message(paste("Error en modelo 2 para", subtype, "-", outcome_name, ":", e$message))
      return(NULL)
    })
    
    # Extraer OR para modelo 1
    if (!is.null(model1)) {
      coefs1 <- coef(summary(model1))
      aqp8_index1 <- which(rownames(coefs1) == "`ΔΔCT AQP8`")
      
      if (length(aqp8_index1) > 0) {
        aqp8_coef1 <- coefs1[aqp8_index1, 1]
        aqp8_se1 <- coefs1[aqp8_index1, 2]
        aqp8_p1 <- coefs1[aqp8_index1, 4]
        
        or1 <- exp(aqp8_coef1)
        lci1 <- exp(aqp8_coef1 - 1.96 * aqp8_se1)
        uci1 <- exp(aqp8_coef1 + 1.96 * aqp8_se1)
        
        or_results <- rbind(or_results, data.frame(
          outcome = outcome_name,
          subtype = subtype,
          model = "Model 1 (Unadjusted)",
          OR = or1,
          LCI = lci1,
          UCI = uci1,
          p_value = aqp8_p1
        ))
      }
    }
    
    # Extraer OR para modelo 2
    if (!is.null(model2)) {
      coefs2 <- coef(summary(model2))
      aqp8_index2 <- which(rownames(coefs2) == "`ΔΔCT AQP8`")
      
      if (length(aqp8_index2) > 0) {
        aqp8_coef2 <- coefs2[aqp8_index2, 1]
        aqp8_se2 <- coefs2[aqp8_index2, 2]
        aqp8_p2 <- coefs2[aqp8_index2, 4]
        
        or2 <- exp(aqp8_coef2)
        lci2 <- exp(aqp8_coef2 - 1.96 * aqp8_se2)
        uci2 <- exp(aqp8_coef2 + 1.96 * aqp8_se2)
        
        or_results <- rbind(or_results, data.frame(
          outcome = outcome_name,
          subtype = subtype,
          model = "Model 2 (Adjusted)",
          OR = or2,
          LCI = lci2,
          UCI = uci2,
          p_value = aqp8_p2
        ))
      }
    }
  }
  
  # Calcular estadísticas descriptivas
  outcome_stats <- data_filtered %>%
    group_by(Subtipoanatóm) %>%
    summarise(
      total_n = n(),
      events = sum(.data[[outcome_var]] == event_level),
      event_rate = sprintf("%.1f%%", 100 * sum(.data[[outcome_var]] == event_level) / n())
    )
  
  # Devolver resultados
  list(
    or_results = or_results,
    outcome_stats = outcome_stats
  )
}

#-------------------------------------------------------------------------------
# 2. EJECUTAR ANÁLISIS PARA CADA OUTCOME
#-------------------------------------------------------------------------------

# Definir los outcomes a analizar con sus niveles correctos
outcomes <- list(
  list(
    var = "Histol_dif", 
    name = "Histological Undifferentiation",
    ref = "Well differentiated",
    event = "Some grade of undifferentiation"
  ),
  list(
    var = "metastasis", 
    name = "Metastasis",
    ref = "No",
    event = "Yes"
  ),
  list(
    var = "Progres", 
    name = "Tumour Progression",
    ref = "No",
    event = "Yes"
  )
)

# Ejecutar análisis para cada outcome
all_results <- lapply(outcomes, function(outcome) {
  analyze_outcome(outcome$var, outcome$name, outcome$ref, outcome$event)
})

# Combinar todos los resultados de OR
combined_or_results <- bind_rows(lapply(all_results, function(x) {
  if (is.null(x$or_results) || nrow(x$or_results) == 0) {
    return(NULL)
  }
  return(x$or_results)
}))

# Combinar estadísticas por outcome
combined_stats <- bind_rows(lapply(1:length(outcomes), function(i) {
  if (is.null(all_results[[i]]$outcome_stats) || nrow(all_results[[i]]$outcome_stats) == 0) {
    return(NULL)
  }
  all_results[[i]]$outcome_stats %>%
    mutate(outcome = outcomes[[i]]$name)
}))

# Verificar si hay datos para continuar
if (nrow(combined_or_results) == 0) {
  stop("No hay suficientes datos para ninguno de los outcomes. Verifica los nombres de las variables y los niveles.")
}

#-------------------------------------------------------------------------------
# 3. CREAR FOREST PLOT UNIFICADO
#-------------------------------------------------------------------------------

# Preparar datos para el gráfico
plot_data <- combined_or_results %>%
  # Crear factores ordenados
  mutate(
    # Ordenar outcomes
    outcome = factor(outcome, levels = c(
      "Histological Undifferentiation",
      "Metastasis",
      "Tumour Progression"
    )),
    # Ordenar subtipos
    subtype = factor(subtype),
    # Ordenar modelos
    model = factor(model, levels = c("Model 1 (Unadjusted)", "Model 2 (Adjusted)")),
    # Crear etiqueta compuesta para el gráfico
    label = paste(subtype, model, sep = " - ")
  )

# Determinar límites para el gráfico
max_value <- max(plot_data$UCI, na.rm = TRUE)
max_limit <- min(max(10, ceiling(max_value * 1.1)), 100)

# Crear forest plot unificado
unified_forest_plot <- ggplot(plot_data, 
                              aes(x = OR, xmin = LCI, xmax = UCI, y = label, 
                                  color = subtype, shape = model)) +
  
  # Línea de referencia en OR=1
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales y puntos
  geom_errorbarh(aes(height = 0.2)) +
  geom_point(size = 4) +
  
  # Facetas por outcome
  facet_grid(outcome ~ ., scales = "free_y", space = "free_y", switch = "y") +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Effect of AQP8 Expression on Multiple Colon Cancer Outcomes",
    subtitle = "Stratified by Anatomical Location",
    y = NULL,
    color = "Anatomical Subtype",
    shape = "Model"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.box = "vertical",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12),
    strip.background = element_rect(fill = "grey90"),
    strip.text = element_text(face = "bold"),
    strip.placement = "outside"
  ) +
  
  # Escalas de color y forma
  scale_color_manual(values = c("Distal colon" = "#3A86FF", 
                                "Proximal colon" = "#FF006E")) +
  scale_shape_manual(values = c("Model 1 (Unadjusted)" = 15, 
                                "Model 2 (Adjusted)" = 18))

#-------------------------------------------------------------------------------
# 4. CREAR FOREST PLOT ALTERNATIVO (POR SUBTIPOS)
#-------------------------------------------------------------------------------

# Un diseño alternativo agrupando por subtipo anatómico
alt_forest_plot <- ggplot(plot_data, 
                          aes(x = OR, xmin = LCI, xmax = UCI, y = model, 
                              color = outcome, shape = outcome)) +
  
  # Línea de referencia en OR=1
  geom_vline(xintercept = 1, linetype = "solid", color = "gray20", linewidth = 0.4) +
  
  # Barras de error horizontales y puntos
  geom_errorbarh(aes(height = 0.2), position = position_dodge(width = 0.5)) +
  geom_point(size = 4, position = position_dodge(width = 0.5)) +
  
  # Facetas por subtipo anatómico
  facet_grid(subtype ~ ., scales = "free_y", space = "free_y", switch = "y") +
  
  # Escala X logarítmica
  scale_x_log10(
    breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50),
    name = "Odds Ratio for ΔΔCT AQP8 (Log Scale)"
  ) +
  coord_cartesian(xlim = c(0.1, max_limit)) +
  
  # Etiquetas
  labs(
    title = "Effect of AQP8 Expression on Colon Cancer Outcomes",
    subtitle = "Comparison across Different Anatomical Locations",
    y = NULL,
    color = "Outcome",
    shape = "Outcome"
  ) +
  
  # Tema
  theme_minimal(base_size = 14) +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),
    plot.background = element_rect(fill = "white", colour = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor.x = element_blank(),
    panel.grid.major.y = element_blank(),
    axis.ticks = element_line(color = "black"),
    axis.line.x = element_line(color = "black", linewidth = 0.4),
    legend.position = "top",
    legend.title = element_text(face = "bold"),
    axis.title.x = element_text(margin = margin(t = 10)),
    plot.title = element_text(face = "bold"),
    plot.subtitle = element_text(face = "italic", size = 12),
    strip.background = element_rect(fill = "grey90"),
    strip.text = element_text(face = "bold"),
    strip.placement = "outside"
  ) +
  
  # Escalas de color y forma
  scale_color_manual(values = c(
    "Histological Undifferentiation" = "#2E86AB", 
    "Metastasis" = "#D13E37",
    "Tumour Progression" = "#6C4F77"
  )) +
  scale_shape_manual(values = c(
    "Histological Undifferentiation" = 15, 
    "Metastasis" = 16,
    "Tumour Progression" = 17
  ))

#-------------------------------------------------------------------------------
# 5. CREAR TABLA UNIFICADA
#-------------------------------------------------------------------------------

# Formatear los resultados para una tabla más completa
formatted_table <- combined_or_results %>%
  mutate(
    OR_CI = sprintf("%.2f (%.2f-%.2f)", OR, LCI, UCI),
    p_value_formatted = ifelse(p_value < 0.001, 
                               "<0.001", 
                               sprintf("%.3f", p_value))
  ) %>%
  left_join(combined_stats, by = c("outcome" = "outcome", "subtype" = "Subtipoanatóm")) %>%
  select(
    Outcome = outcome,
    `Anatomical Subtype` = subtype, 
    Model = model, 
    `Sample Size` = total_n,
    `Events (%)` = event_rate,
    `OR (95% CI)` = OR_CI, 
    `p-value` = p_value_formatted
  ) %>%
  arrange(Outcome, `Anatomical Subtype`, Model)

# Mostrar la tabla formateada
print("Tabla unificada de resultados para todos los outcomes:")
print(formatted_table, row.names = FALSE)

# Guardar la tabla como archivo CSV (opcional)
write.csv(formatted_table, "AQP8_unified_outcomes_table.csv", row.names = FALSE)

#-------------------------------------------------------------------------------
# 6. GUARDAR GRÁFICOS (opcional)
#-------------------------------------------------------------------------------

# Guardar el forest plot principal
ggsave("AQP8_unified_forest_plot.png", unified_forest_plot, 
       width = 10, height = 8, dpi = 300)

# Guardar el forest plot alternativo
ggsave("AQP8_alternative_forest_plot.png", alt_forest_plot, 
       width = 10, height = 8, dpi = 300)

#-------------------------------------------------------------------------------
# 7. MOSTRAR LOS GRÁFICOS
#-------------------------------------------------------------------------------
print(unified_forest_plot)
print(alt_forest_plot)

# ============================================================
# 05_taules_cap4.R - Taules del capitol 4 de la memoria (BD v26)
# TFM Diablo Wasi. Executar des d'analysis/:
#   Rscript scripts/05_taules_cap4.R
# Requereix nomes data/matriu_AX_cinc_valors.csv (derivada versionada,
# regenerable amb 00+01) i R base: cap paquet addicional.
# Poblacio (D1), codificacio (D2) i clusters (D4): identics a 02-04.
#  1 TAULA-4.1: corpus per sector x tipologia (106 EA i poblacio de 92)
#  2 Riquesa constructiva per tipologia (mateixa variable que la passada 2,
#    analisi 1): descriptius, Kruskal-Wallis i parells amb correccio de
#    Holm; taula per EA per a la figura; control per familia constructiva
# ============================================================
set.seed(26)
for (d in c("output/logs","output/tables"))
  dir.create(d, recursive = TRUE, showWarnings = FALSE)
sink("output/logs/resultats_pas3.txt", split = TRUE)
cat("=== PASSADA 3: TAULES DEL CAPITOL 4 ===", format(Sys.time()), "\n")
cat("Font: data/matriu_AX_cinc_valors.csv (T_STRUCTURES v26 x L_SECTORS x",
    "L_TYPOLOGY, via 01_build_matriu_AX.py)\n\n")

M  <- read.csv("data/matriu_AX_cinc_valors.csv", check.names = FALSE)
ll <- c("B","C","D","E","F","G","I","J","K","L","M","N","O","Q","R","S","T","V","X","Z")

# Ordre de presentacio de les tipologies (xifrari del TFM, bloc Corpus)
tip_ordre <- c("EA-MAU Mausoleum/Chullpa", "EA-CAM Funerary Chamber",
               "EA-TER Ledge Terrace", "EA-PLA-V Aerial Platform",
               "Unclassifiable", "NIX Natural Niche", "CAV Cave/Cavern",
               "MEN Isolated Structural Element", "PR Rock Art")
tip_curt <- c("MAU","CAM","TER","PLA-V","NC","NIX","CAV","MEN","PR")
names(tip_curt) <- tip_ordre
stopifnot(all(M$Typology %in% tip_ordre))
M$Sector_curt <- substr(M$Code, 1, 6)          # DW-S01 ... DW-S06 (del codi d'EA)
M$Typology    <- factor(M$Typology, levels = tip_ordre)

# Taula de contingencia amb totals de fila i columna, en CSV i en Markdown
md_taula <- function(tt, etiq_col = colnames(tt)) {
  cat("| Sector | ", paste(etiq_col, collapse = " | "), " |\n", sep = "")
  cat("|---|", paste(rep("---:", ncol(tt)), collapse = "|"), "|\n", sep = "")
  for (i in seq_len(nrow(tt)))
    cat("| ", rownames(tt)[i], " | ", paste(tt[i, ], collapse = " | "), " |\n", sep = "")
}
escriu_contingencia <- function(df, fitxer, titol) {
  tt <- table(df$Sector_curt, droplevels(df$Typology))
  tt <- addmargins(tt, FUN = sum, quiet = TRUE)
  rownames(tt)[nrow(tt)] <- "Total"; colnames(tt)[ncol(tt)] <- "Total"
  out <- cbind(Sector = rownames(tt), as.data.frame.matrix(tt))
  write.csv(out, fitxer, row.names = FALSE)
  cat("--- ", titol, " (n = ", nrow(df), ") ---\n", sep = "")
  print(tt)
  cat("\nMarkdown:\n\n")
  md_taula(tt, c(tip_curt[colnames(tt)[-ncol(tt)]], "Total"))
  cat("\nFitxer:", fitxer, "\n\n")
  invisible(tt)
}

# ---------- 1. TAULA-4.1: CORPUS PER SECTOR x TIPOLOGIA ----------
cat("=== 1. TAULA-4.1: corpus per sector x tipologia ===\n\n")
t106 <- escriu_contingencia(M, "output/tables/taula_4_1_sector_x_tipologia_106.csv",
                            "Corpus complet (9 tipologies)")

# D1: poblacio analitica (mateix criteri que 02-04)
classe_ok <- M$Record_Class %in% c("Built funerary structure","Natural funerary context")
cobertura <- rowSums(!is.na(M[, ll]))
pobl <- classe_ok & cobertura >= ceiling(length(ll)/2)
A <- M[pobl, ]
cat("Decisio D1: excloses", sum(!classe_ok), "EA per classe de registre (",
    paste(sort(M$Code[!classe_ok]), collapse = ", "), ")\n")
cat("            i", sum(classe_ok & !pobl), "per cobertura < 50 % (",
    paste(sort(M$Code[classe_ok & !pobl]), collapse = ", "), ")\n\n")
t92 <- escriu_contingencia(A, "output/tables/taula_4_1_sector_x_tipologia_92.csv",
                           "Poblacio analitica (7 tipologies, decisio D1)")

# Comprovacio contra els marginals del xifrari v05 (bloc Corpus)
esperat <- list(
  s106 = c(`DW-S01`=71,`DW-S02`=5,`DW-S03`=4,`DW-S04`=20,`DW-S05`=2,`DW-S06`=4),
  t106 = c(MAU=17,CAM=14,TER=19,`PLA-V`=4,NC=9,NIX=19,CAV=16,MEN=4,PR=4),
  s92  = c(`DW-S01`=60,`DW-S02`=5,`DW-S03`=4,`DW-S04`=18,`DW-S05`=2,`DW-S06`=3),
  t92  = c(MAU=17,CAM=14,TER=18,`PLA-V`=3,NC=7,NIX=18,CAV=15))
obtingut <- list(
  s106 = t106[rownames(t106) != "Total", "Total"],
  t106 = setNames(t106["Total", colnames(t106) != "Total"], tip_curt[colnames(t106)[colnames(t106) != "Total"]]),
  s92  = t92[rownames(t92) != "Total", "Total"],
  t92  = setNames(t92["Total", colnames(t92) != "Total"], tip_curt[colnames(t92)[colnames(t92) != "Total"]]))
quadra <- sapply(names(esperat), function(k)
  identical(as.integer(obtingut[[k]][names(esperat[[k]])]), as.integer(esperat[[k]])))
cat("Comprovacio dels marginals contra el xifrari v05:",
    if (all(quadra)) "tots quadren" else paste("NO quadren:", paste(names(quadra)[!quadra], collapse = ", ")),
    "\n\n")
stopifnot(all(quadra))

# ---------- 2. RIQUESA CONSTRUCTIVA PER TIPOLOGIA ----------
cat("=== 2. Riquesa constructiva per tipologia (n = ", nrow(A), ") ===\n", sep = "")
cat("Riquesa = elements A-X construits (valor 1 en la codificacio D2) per EA;",
    "identica a l'analisi 1 de la passada 2.\n\n")
X <- as.matrix(A[, ll])
P <- ifelse(is.na(X), NA, ifelse(X %in% c(1,2,3), 1, ifelse(X == 0, 0, NA)))
P <- matrix(as.numeric(P), nrow = nrow(X), dimnames = list(A$Code, ll))
A$Riquesa <- rowSums(P == 1, na.rm = TRUE)
A$Typology <- droplevels(A$Typology)

# Familia constructiva: mateixa cadena deterministica que 02-04 (Jaccard + UPGMA, k = 5)
jac <- function(P) { n <- nrow(P)
  D <- matrix(NA_real_, n, n, dimnames = list(rownames(P), rownames(P)))
  for (i in 1:n) for (j in i:n) {
    ok <- !is.na(P[i,]) & !is.na(P[j,])
    a <- sum(P[i,ok]==1 & P[j,ok]==1); b <- sum(P[i,ok]!=P[j,ok]); u <- a+b
    D[i,j] <- D[j,i] <- if (u==0) 0 else 1 - a/u }
  as.dist(D) }
A$Familia <- factor(cutree(hclust(jac(P), method = "average"), 5)[A$Code])

cat("Resum global de la riquesa (ha de coincidir amb la passada 2):\n")
print(summary(A$Riquesa))
cat("\n")

per_EA <- data.frame(Code = A$Code, Sector = A$Sector_curt,
                     Tipologia = as.character(A$Typology),
                     Familia = as.integer(as.character(A$Familia)),
                     Riquesa = as.integer(A$Riquesa))
per_EA <- per_EA[order(per_EA$Code), ]
write.csv(per_EA, "output/tables/riquesa_per_EA.csv", row.names = FALSE)
cat("Fitxer per EA: output/tables/riquesa_per_EA.csv (", nrow(per_EA), " files)\n\n", sep = "")

resum_grup <- function(x, g) {
  r <- t(sapply(split(x, g), function(v)
    c(n = length(v), Min = min(v), Q1 = unname(quantile(v, .25)),
      Mediana = median(v), Q3 = unname(quantile(v, .75)), Max = max(v),
      Mitjana = round(mean(v), 2))))
  as.data.frame(r)
}
md_resum <- function(r, etiq) {
  cat("| ", etiq, " | n | Min | Q1 | Mediana | Q3 | Max | Mitjana |\n", sep = "")
  cat("|---|---:|---:|---:|---:|---:|---:|---:|\n")
  for (i in seq_len(nrow(r)))
    cat("| ", rownames(r)[i], " | ", paste(unlist(r[i, ]), collapse = " | "), " |\n", sep = "")
}
# Prova global + parells (Wilcoxon-Mann-Whitney, aprox. normal, Holm)
contrast <- function(x, g, nom) {
  kt <- kruskal.test(x ~ g)
  cat(sprintf("Kruskal-Wallis (%s): H = %.2f, gl = %d, p = %s\n", nom,
              kt$statistic, kt$parameter, format(signif(kt$p.value, 3))))
  if (kt$p.value >= 0.05) { cat("No significatiu: no es fan comparacions per parells.\n\n"); return(invisible(NULL)) }
  pw <- pairwise.wilcox.test(x, g, p.adjust.method = "holm", exact = FALSE)
  cat("\nComparacions per parells (Wilcoxon-Mann-Whitney, p ajustat per Holm):\n")
  print(round(pw$p.value, 4), na.print = "")
  nn <- table(g)
  cat("\nLectura (alfa = 0,05; les diferencies s'entenen sobre la mediana i",
      "el rang de la riquesa):\n")
  nivells <- levels(g)
  m <- pw$p.value
  p_parell <- function(a, b) {   # la matriu de pairwise.wilcox.test es triangular
    v <- NA_real_
    if (b %in% rownames(m) && a %in% colnames(m)) v <- m[b, a]
    if (is.na(v) && a %in% rownames(m) && b %in% colnames(m)) v <- m[a, b]
    v }
  for (a in nivells) {
    p_a <- sapply(setdiff(nivells, a), function(b) p_parell(a, b))
    dif <- names(p_a)[!is.na(p_a) & p_a < 0.05]
    nod <- names(p_a)[!is.na(p_a) & p_a >= 0.05]
    cat(sprintf("- %s (n = %d, mediana %s): es distingeix de %s; no es distingeix de %s.\n",
        a, nn[a], format(median(x[g == a])),
        if (length(dif)) paste(dif, collapse = ", ") else "cap",
        if (length(nod)) paste(nod, collapse = ", ") else "cap"))
  }
  petits <- names(nn)[nn < 5]
  if (length(petits))
    cat("Nota: els grups amb n < 5 (", paste(petits, collapse = ", "),
        ") tenen poca potencia; les seues comparacions son nomes indicatives.\n", sep = "")
  cat("\n")
  invisible(pw)
}

cat("--- 2a. Per tipologia ---\n")
rt <- resum_grup(A$Riquesa, A$Typology)
print(rt)
write.csv(cbind(Tipologia = rownames(rt), rt),
          "output/tables/taula_riquesa_tipologia.csv", row.names = FALSE)
cat("\nMarkdown:\n\n"); md_resum(rt, "Tipologia")
cat("\nFitxer: output/tables/taula_riquesa_tipologia.csv\n\n")
contrast(A$Riquesa, A$Typology, "riquesa x tipologia")

cat("--- 2b. Control: per familia constructiva (cluster de la passada 1) ---\n")
rf <- resum_grup(A$Riquesa, A$Familia)
print(rf)
write.csv(cbind(Familia = rownames(rf), rf),
          "output/tables/taula_riquesa_familia.csv", row.names = FALSE)
cat("\nMarkdown:\n\n"); md_resum(rf, "Familia")
cat("\nFitxer: output/tables/taula_riquesa_familia.csv\n\n")
contrast(A$Riquesa, A$Familia, "riquesa x familia")

cat("Tipologia x familia (n = ", nrow(A), "):\n", sep = "")
print(table(A$Typology, A$Familia))
cat("\nLa figura fig_riquesa_tipologia.pdf la genera la passada 2 (03) a partir",
    "de la mateixa variable; riquesa_per_EA.csv permet redibuixar-la.\n")
sink()
cat("FET passada 3.\n")

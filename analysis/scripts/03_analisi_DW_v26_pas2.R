# ============================================================
# analisi_DW_v26_pas2.R - Passada 2 (analisis 1-5 firmades 2026-09-07)
# 1 Riquesa constructiva x suport/altitud/qualitat + esforc (murs v5)
# 2 MNI x Area_m2 (proxy; Volume_m3 buit al corpus - troballa documental)
# 3 Punts d'articulacio del graf de connexions (OE3, circulacio)
# 4 Afectacio/supervivencia per element (biaix tafonomic, fusta vs pedra)
# 5 Decoracio x clusters i x posicio (H04)
# ============================================================
suppressPackageStartupMessages({ library(ggplot2); library(igraph) })
set.seed(26)
for (d in c("output/figures","output/logs","output/tables")) dir.create(d, recursive = TRUE, showWarnings = FALSE)
sink("output/logs/resultats_pas2.txt", split = TRUE)
cat("=== PASSADA 2 ===", format(Sys.time()), "\n\n")

M  <- read.csv("data/matriu_AX_cinc_valors.csv", check.names = FALSE)
TS <- read.csv("data/T_STRUCTURES.csv")
CN <- read.csv("data/T_CONNECTIONS.csv")
DE <- read.csv("data/T_DECORATIONS.csv")
LB <- read.csv("data/L_STRUCT_BODY.csv")
LD <- read.csv("data/L_DEC_TYPE.csv")
LE <- read.csv("data/T_LOST_ELEMENTS.csv")
MU <- read.csv("data/metriques_murs_v5.csv", fileEncoding = "UTF-8-BOM")

lletres <- c("B","C","D","E","F","G","I","J","K","L","M","N","O","Q","R","S","T","V","X","Z")
pobl <- M$Record_Class %in% c("Built funerary structure","Natural funerary context") &
        rowSums(!is.na(M[, lletres])) >= 10
A <- M[pobl, ]
X <- as.matrix(A[, lletres])

# Reproduir clusters de la passada 1 (mateixa cadena deterministica)
P <- ifelse(is.na(X), NA, ifelse(X %in% c(1,2,3), 1, ifelse(X == 0, 0, NA)))
P <- matrix(as.numeric(P), nrow = nrow(X), dimnames = list(A$Code, lletres))
jaccard_na <- function(P) { n <- nrow(P); D <- matrix(NA_real_, n, n, dimnames = list(rownames(P), rownames(P)))
  for (i in 1:n) for (j in i:n) { ok <- !is.na(P[i,]) & !is.na(P[j,])
    a <- sum(P[i,ok]==1 & P[j,ok]==1); b <- sum(P[i,ok]!=P[j,ok]); u <- a+b
    D[i,j] <- D[j,i] <- if (u==0) 0 else 1 - a/u }; as.dist(D) }
A$Cluster <- factor(cutree(hclust(jaccard_na(P), method="average"), 5)[A$Code])

# ---------- 1. RIQUESA CONSTRUCTIVA ----------
A$Riquesa <- rowSums(P == 1, na.rm = TRUE)
A$Avaluables <- rowSums(!is.na(P))
T2 <- TS[match(A$Code, TS$Code), ]
cat("=== 1. Riquesa constructiva (elements construits per EA) ===\n")
print(summary(A$Riquesa))
sp <- function(x, y, nom) { ok <- !is.na(x) & !is.na(y)
  if (sum(ok) < 5) { cat(nom, ": n insuficient (", sum(ok), ")\n"); return(invisible()) }
  ct <- suppressWarnings(cor.test(x[ok], y[ok], method = "spearman"))
  cat(sprintf("%-28s n=%3d  rho=%+.3f  p=%s\n", nom, sum(ok),
      ct$estimate, format(signif(ct$p.value,3)))) }
sp(A$Riquesa, T2$Support_Width_m,     "Riquesa x Support_Width_m")
sp(A$Riquesa, T2$Support_Depth_m,     "Riquesa x Support_Depth_m")
sp(A$Riquesa, T2$Altitude_masl,       "Riquesa x Altitude_masl")
sp(A$Riquesa, T2$Height_Above_Base_m, "Riquesa x Height_Above_Base_m")
mq <- factor(T2$Masonry_Quality, levels = c("Poor","Moderate","Good"))
ok <- !is.na(mq)
if (sum(ok) > 10) { kt <- kruskal.test(A$Riquesa[ok] ~ mq[ok])
  cat("Riquesa x Masonry_Quality (Kruskal-Wallis): H =", round(kt$statistic,2),
      " p =", signif(kt$p.value,3), "\n")
  print(tapply(A$Riquesa[ok], mq[ok], median)) }
# Esforc: agregat de murs v5 per estructura
names(MU)[1] <- "Code"
ef <- aggregate(cbind(Area_Cons_m2, Volume_Prism_m3) ~ Code, MU, sum, na.rm = TRUE)
A2 <- merge(A, ef, by = "Code")
cat("\nEA amb metriques de mur agregades:", nrow(A2), "\n")
sp(A2$Riquesa, A2$Area_Cons_m2,    "Riquesa x Area murs (m2)")
sp(A2$Riquesa, A2$Volume_Prism_m3, "Riquesa x Volum prisma (m3)")
ggsave("output/figures/fig_riquesa_tipologia.pdf", width = 9, height = 5,
  plot = ggplot(A, aes(reorder(Typology, Riquesa, median), Riquesa)) +
    geom_boxplot(fill = "grey85", outlier.shape = 1) +
    geom_jitter(width = .15, alpha = .4, size = 1.2) +
    coord_flip() + labs(title = "Riquesa constructiva per tipologia",
      subtitle = "Elements A-X construits complets (valor 1) per estructura; n = 92",
      x = NULL, y = "Elements construits") + theme_minimal(base_size = 11))

# ---------- 2. MNI x AREA ----------
cat("\n=== 2. MNI x Area_m2 (proxy d'espai funerari; Volume_m3 buit) ===\n")
mn <- TS[!is.na(TS$MNI) & !is.na(TS$Area_m2), c("Code","MNI","Area_m2","Looting")]
cat("n =", nrow(mn), "\n")
ct <- suppressWarnings(cor.test(mn$MNI, mn$Area_m2, method = "spearman"))
cat("Spearman rho =", round(ct$estimate,3), " p =", signif(ct$p.value,3), "\n")
mn$Espoli <- factor(ifelse(mn$Looting == 1, "Espoliada",
                    ifelse(mn$Looting == 0, "No espoliada", "No determinable")))
ggsave("output/figures/fig_mni_area.pdf", width = 7, height = 5.5,
  plot = ggplot(mn, aes(Area_m2, MNI, shape = Espoli)) +
    geom_point(size = 2.6) + scale_shape_manual(values = c(16, 1, 4)) +
    geom_text(data = subset(mn, MNI > 10), aes(label = sub("DW-","",Code)),
              vjust = -0.9, size = 2.8) +
    labs(title = "MNI vs superficie funeraria (n = 13)",
         subtitle = sprintf("Spearman rho = %.2f, p = %.3f - lectura tafonomica obligada",
                            ct$estimate, ct$p.value),
         x = "Area funeraria (m2)", y = "MNI") + theme_minimal(base_size = 11))

# ---------- 3. PUNTS D'ARTICULACIO ----------
cat("\n=== 3. Graf: punts d'articulacio (OE3) ===\n")
id2code <- setNames(TS$Code, TS$ID)
g <- graph_from_data_frame(data.frame(from = id2code[as.character(CN$ID_Struct_A)],
                                      to   = id2code[as.character(CN$ID_Struct_B)]),
      directed = FALSE, vertices = data.frame(name = TS$Code))
ap <- articulation_points(g)
cat("Punts d'articulacio:", length(ap), "\n")
apdf <- data.frame(Code = names(ap),
                   Grau = degree(g)[names(ap)],
                   Tipologia = NA_character_)
apdf$Tipologia <- M$Typology[match(apdf$Code, M$Code)]
apdf <- apdf[order(-apdf$Grau), ]
print(apdf, row.names = FALSE)
write.csv(apdf, "output/tables/taula_punts_articulacio.csv", row.names = FALSE)
pdf("output/figures/fig_graf_articulacio.pdf", width = 10, height = 8)
sub <- induced_subgraph(g, V(g)[degree(g) > 0])
V(sub)$col <- ifelse(V(sub)$name %in% names(ap), "black", "grey80")
V(sub)$sz  <- ifelse(V(sub)$name %in% names(ap), 8, 5)
set.seed(26)
plot(sub, vertex.color = V(sub)$col, vertex.size = V(sub)$sz,
     vertex.label.cex = .5, vertex.label.dist = .6, edge.color = "grey55",
     layout = layout_with_fr(sub),
     main = "Punts d'articulacio de la xarxa de connexions (negre = nus critic)")
dev.off()

# ---------- 4. AFECTACIO PER ELEMENT ----------
cat("\n=== 4. Afectacio per element (2 parcial + 3 perdut sobre construit) ===\n")
comp <- t(sapply(lletres, function(l) { v <- X[, l]
  c(complet = sum(v==1, na.rm=TRUE), parcial = sum(v==2, na.rm=TRUE),
    perdut = sum(v==3, na.rm=TRUE)) }))
comp <- as.data.frame(comp)
comp$constr <- rowSums(comp)
comp$afectacio <- with(comp, ifelse(constr==0, NA, (parcial+perdut)/constr))
comp$fusta <- rownames(comp) %in% c("E","F","S")
comp <- comp[order(-comp$afectacio), ]
print(round(comp[comp$constr>0, c("complet","parcial","perdut","constr","afectacio")], 2))
le_n <- table(LE$Element_Code)
cat("\nEvidencies a T_LOST_ELEMENTS per element:\n"); print(le_n)
cat("\nAfectacio mitjana ponderada fusta (E,F,S):",
    round(weighted.mean(comp$afectacio[comp$fusta], comp$constr[comp$fusta], na.rm=TRUE),2),
    "| petris:",
    round(weighted.mean(comp$afectacio[!comp$fusta], comp$constr[!comp$fusta], na.rm=TRUE),2), "\n")
pd <- data.frame(Element = rep(rownames(comp), 3),
  Estat = rep(c("Complet","Parcial","Perdut atestat"), each = nrow(comp)),
  n = c(comp$complet, comp$parcial, comp$perdut))
pd <- pd[pd$n > 0 | pd$Estat == "Complet", ]
pd$Element <- factor(pd$Element, levels = rownames(comp))
pd$Estat <- factor(pd$Estat, levels = c("Perdut atestat","Parcial","Complet"))
ggsave("output/figures/fig_afectacio_elements.pdf", width = 9.5, height = 5,
  plot = ggplot(pd, aes(Element, n, fill = Estat)) +
    geom_col(color = "white", linewidth = .2) +
    scale_fill_manual(values = c("grey15","grey55","grey85")) +
    labs(title = "Estat dels elements construits (n = 92 EA)",
      subtitle = "Ordenats per afectacio (parcial + perdut / construit); E, F i S son portadors de fusta",
      x = "Element A-X", y = "Ocurrencies") + theme_minimal(base_size = 11))

# ---------- 5. DECORACIO ----------
cat("\n=== 5. Decoracio (T_DECORATIONS, 84 files) ===\n")
DE$Body <- LB$Name[match(DE$ID_Struct_Body, LB$ID)]
DE$Tipus <- LD$Name[match(DE$ID_Dec_Type, LD$ID)]
DE$Code <- id2code[as.character(DE$ID_Structure)]
cat("Files:", nrow(DE), "| Estructures decorades:", length(unique(DE$Code)), "\n")
cat("\nPer tipus:\n"); print(sort(table(DE$Tipus), decreasing = TRUE))
cat("\nPer cos posicional:\n"); print(sort(table(DE$Body), decreasing = TRUE))
A$Decorada <- ifelse(A$Code %in% DE$Code, 1, 0)
dt <- table(A$Cluster, A$Decorada); colnames(dt) <- c("Sense","Amb")
cat("\nDecoracio x Cluster:\n"); print(dt)
ftest <- fisher.test(dt, simulate.p.value = TRUE, B = 20000)
cat("Fisher p =", signif(ftest$p.value, 3), "\n")
tt <- table(A$Typology, A$Decorada); colnames(tt) <- c("Sense","Amb")
cat("\nDecoracio x Tipologia:\n"); print(tt)
pdc <- as.data.frame(sort(table(DE$Body), decreasing = TRUE))
names(pdc) <- c("Cos","n")
ggsave("output/figures/fig_decoracio_posicio.pdf", width = 8, height = 4.5,
  plot = ggplot(pdc, aes(reorder(Cos, n), n)) +
    geom_col(fill = "grey35") + coord_flip() +
    labs(title = "Distribucio posicional de la decoracio (84 registres)",
         x = NULL, y = "Registres") + theme_minimal(base_size = 11))
sink()
cat("FET passada 2.\n")

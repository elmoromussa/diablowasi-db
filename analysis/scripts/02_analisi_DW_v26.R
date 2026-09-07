# ============================================================
# analisi_DW_v26.R - Analisi del corpus de Diablo Wasi (BD v26)
# TFM Esteve Ribera-Torro (MAGIP, UA) - corpus congelat v26
# Decisions D1-D4 tancades en sessio 2026-09-07:
#  D1 poblacio: Built funerary structure + Natural funerary context;
#     exclosos PR i MEN; llindar de cobertura >=50% elements avaluables
#  D2 codificacio primaria "construit": {1,2,3}->1, {0}->0, {9}->NA;
#     sensibilitat "sobreviu": {1,2}->1, {0,3}->0, {9}->NA
#  D3 elements: 21 del QRY_13 menys A (constant-absent); Support_Modified fora
#  D4 metodes: Jaccard + UPGMA (estructures), Jaccard element x element,
#     contrast clusters vs tipologia i sector; igraph sobre T_CONNECTIONS;
#     rosa d'orientacions sobre Facade_Azimuth_Deg
# ============================================================

suppressPackageStartupMessages({ library(ggplot2); library(igraph) })
set.seed(26)
for (d in c("output/figures","output/logs","output/tables")) dir.create(d, recursive = TRUE, showWarnings = FALSE)
sink("output/logs/resultats_analisi.txt", split = TRUE)
cat("=== ANALISI DW v26 ===", format(Sys.time()), "\n\n")

M  <- read.csv("data/matriu_AX_cinc_valors.csv", check.names = FALSE)
TS <- read.csv("data/T_STRUCTURES.csv")
CN <- read.csv("data/T_CONNECTIONS.csv")

lletres_tot <- c("B","C","D","E","F","G","I","J","K","L","M","N","O","Q","R","S","T","V","X","Z")

# ---------- D1: poblacio analitica ----------
pobl <- M$Record_Class %in% c("Built funerary structure","Natural funerary context")
cob  <- rowSums(!is.na(M[, lletres_tot]))
prou <- cob >= ceiling(length(lletres_tot)/2)
A <- M[pobl & prou, ]
cat("Poblacio analitica:", nrow(A), "EA (excloses:",
    sum(pobl & !prou), "per cobertura <50%,", sum(!pobl), "per Record_Class)\n\n")

# ---------- D2: codificacions ----------
codifica <- function(df, mode) {
  X <- as.matrix(df[, lletres_tot])
  if (mode == "construit") { P <- ifelse(is.na(X), NA, ifelse(X %in% c(1,2,3), 1, ifelse(X == 0, 0, NA))) }
  else                     { P <- ifelse(is.na(X), NA, ifelse(X %in% c(1,2),   1, ifelse(X %in% c(0,3), 0, NA))) }
  P <- matrix(as.numeric(P), nrow = nrow(X), dimnames = list(df$Code, lletres_tot))
  P
}
Pc <- codifica(A, "construit")
Ps <- codifica(A, "sobreviu")

# ---------- Jaccard amb NA per parells ----------
jaccard_na <- function(P) {
  n <- nrow(P); D <- matrix(NA_real_, n, n, dimnames = list(rownames(P), rownames(P)))
  for (i in 1:n) for (j in i:n) {
    ok <- !is.na(P[i,]) & !is.na(P[j,])
    a <- sum(P[i,ok]==1 & P[j,ok]==1); b <- sum(P[i,ok]!=P[j,ok]); u <- a + b
    D[i,j] <- D[j,i] <- if (u == 0) 0 else 1 - a/u   # cap 1 compartit i cap desacord -> dist 0
  }
  as.dist(D)
}
Dc <- jaccard_na(Pc)

# ---------- Clustering d'estructures (UPGMA) ----------
hc <- hclust(Dc, method = "average")
sil_mitjana <- function(D, cl) {
  Dm <- as.matrix(D); ks <- unique(cl); if (length(ks) < 2) return(NA)
  s <- sapply(seq_along(cl), function(i) {
    ai <- mean(Dm[i, cl == cl[i] & seq_along(cl) != i])
    bi <- min(sapply(setdiff(ks, cl[i]), function(k) mean(Dm[i, cl == k])))
    if (is.nan(ai)) 0 else (bi - ai)/max(ai, bi) })
  mean(s)
}
sils <- sapply(2:8, function(k) sil_mitjana(Dc, cutree(hc, k)))
names(sils) <- 2:8
cat("Silueta mitjana per k:\n"); print(round(sils, 3))
k_opt <- as.integer(names(which.max(sils)))
cat("k optim (silueta):", k_opt, "\n\n")
cl <- cutree(hc, k_opt)
A$Cluster <- factor(cl[A$Code])

# Estabilitat tafonomica: mateixos talls amb la codificacio "sobreviu"
cl_s <- cutree(hclust(jaccard_na(Ps), method = "average"), k_opt)
concord <- mean(sapply(unique(cl), function(k) {
  membres <- names(cl)[cl == k]
  max(table(cl_s[membres]))/length(membres) }))
cat("Sensibilitat D2 (construit vs sobreviu): concordanca mitjana de pertinenca =",
    round(concord, 3), "\n\n")

# Dendrograma
pdf("output/figures/fig_dendrograma_UPGMA.pdf", width = 11, height = 6)
par(mar = c(6,4,3,1))
plot(hc, hang = -1, cex = 0.45, xlab = "", sub = "",
     main = sprintf("Clustering UPGMA (dist. Jaccard, codificacio 'construit') - k = %d", k_opt))
rect.hclust(hc, k = k_opt, border = "grey40")
dev.off()

# ---------- Contrastos H01/H02 ----------
cat("=== Cluster x Tipologia ===\n")
tt <- table(A$Cluster, A$Typology); print(tt)
ft <- chisq.test(tt, simulate.p.value = TRUE, B = 20000)
cat("Chi2 Monte Carlo: X2 =", round(ft$statistic,2), " p =", signif(ft$p.value,3), "\n\n")
cat("=== Cluster x Sector ===\n")
ts <- table(A$Cluster, A$Sector); print(ts)
fs <- chisq.test(ts, simulate.p.value = TRUE, B = 20000)
cat("Chi2 Monte Carlo: X2 =", round(fs$statistic,2), " p =", signif(fs$p.value,3), "\n\n")

# Perfil dels clusters: prevalenca de cada element per cluster
prev <- t(sapply(levels(A$Cluster), function(k)
  colMeans(Pc[A$Code[A$Cluster == k], , drop = FALSE], na.rm = TRUE)))
write.csv(round(prev, 3), "output/tables/taula_perfil_clusters.csv")
pd <- data.frame(Cluster = rep(rownames(prev), ncol(prev)),
                 Element = rep(colnames(prev), each = nrow(prev)),
                 Prevalenca = as.vector(prev))
pd$Element <- factor(pd$Element, levels = lletres_tot)
ggsave("output/figures/fig_perfil_clusters.pdf", width = 10, height = 4.5,
  plot = ggplot(pd, aes(Element, Cluster, fill = Prevalenca)) + geom_tile(color = "white") +
    geom_text(aes(label = ifelse(Prevalenca > 0, sprintf("%.2f", Prevalenca), "")), size = 2.6) +
    scale_fill_gradient(low = "white", high = "grey20", limits = c(0,1)) +
    labs(title = "Prevalenca dels elements A-X per cluster (codificacio 'construit')",
         x = "Element", y = "Cluster") + theme_minimal(base_size = 11))

# ---------- Co-ocurrencia element x element ----------
jac_el <- matrix(NA_real_, length(lletres_tot), length(lletres_tot),
                 dimnames = list(lletres_tot, lletres_tot))
for (x in lletres_tot) for (y in lletres_tot) {
  ok <- !is.na(Pc[,x]) & !is.na(Pc[,y])
  a <- sum(Pc[ok,x]==1 & Pc[ok,y]==1); u <- sum(Pc[ok,x]==1 | Pc[ok,y]==1)
  jac_el[x,y] <- if (u == 0) NA else a/u
}
write.csv(round(jac_el, 3), "output/tables/taula_jaccard_elements.csv")
ed <- as.data.frame(as.table(jac_el)); names(ed) <- c("E1","E2","J")
ggsave("output/figures/fig_coocurrencia_elements.pdf", width = 8.5, height = 7.5,
  plot = ggplot(ed, aes(E1, E2, fill = J)) + geom_tile(color = "white") +
    geom_text(aes(label = ifelse(!is.na(J) & J > 0 & E1 != E2, sprintf("%.2f", J), "")), size = 2.2) +
    scale_fill_gradient(low = "white", high = "grey15", na.value = "grey92", limits = c(0,1)) +
    labs(title = "Co-ocurrencia entre elements (index de Jaccard, 'construit')",
         x = NULL, y = NULL, fill = "Jaccard") + theme_minimal(base_size = 11))

# Parelles mes fortes (fora de la diagonal)
up <- which(upper.tri(jac_el), arr.ind = TRUE)
top <- data.frame(E1 = rownames(jac_el)[up[,1]], E2 = colnames(jac_el)[up[,2]],
                  J = jac_el[up])
top <- top[order(-top$J), ]; top <- top[!is.na(top$J) & top$J > 0, ]
cat("=== Top 12 parelles d'elements per Jaccard ===\n"); print(head(top, 12), row.names = FALSE)

# ---------- Graf de connexions (igraph) ----------
cat("\n=== T_CONNECTIONS: families ===\n")
print(table(CN$Connection_Type))
id2code <- setNames(TS$Code, TS$ID)
fisiques <- CN[!grepl("Sequential", CN$Connection_Type, ignore.case = TRUE), ]
g <- graph_from_data_frame(
  data.frame(from = id2code[as.character(fisiques$ID_Struct_A)],
             to   = id2code[as.character(fisiques$ID_Struct_B)],
             tipus = fisiques$Connection_Type),
  directed = FALSE,
  vertices = data.frame(name = TS$Code, sector = TS$ID_Sector))
comp <- components(g)
mides <- sort(table(comp$membership), decreasing = TRUE)
cat("\nComponents connexos (arestes fisiques):", comp$no,
    "| aillats:", sum(comp$csize == 1),
    "| mides dels grups >1:", paste(mides[mides > 1], collapse = ", "), "\n")
grups <- data.frame(Code = names(comp$membership), Component = comp$membership)
write.csv(grups, "output/tables/taula_components_connexions.csv", row.names = FALSE)
pdf("output/figures/fig_graf_connexions.pdf", width = 10, height = 8)
set.seed(26)
sub <- induced_subgraph(g, V(g)[degree(g) > 0])
plot(sub, vertex.size = 6, vertex.label.cex = 0.5, vertex.label.dist = 0.6,
     vertex.color = "grey75", edge.color = "grey40",
     layout = layout_with_fr(sub),
     main = "Graf de connexions fisiques entre estructures (families no sequencials)")
dev.off()

# ---------- Orientacions ----------
az <- TS$Facade_Azimuth_Deg[!is.na(TS$Facade_Azimuth_Deg)]
cat("\n=== Orientacions (Facade_Azimuth_Deg) ===\n")
cat("n =", length(az), "\n")
rad <- az * pi/180
Rbar <- sqrt(mean(cos(rad))^2 + mean(sin(rad))^2)
mitja <- (atan2(mean(sin(rad)), mean(cos(rad))) * 180/pi) %% 360
cat("Mitjana circular =", round(mitja,1), "graus | R =", round(Rbar,3),
    "| test de Rayleigh z =", round(length(az)*Rbar^2, 2),
    " p ~", signif(exp(-length(az)*Rbar^2), 3), "\n")
od <- data.frame(az = az)
ggsave("output/figures/fig_rosa_orientacions.pdf", width = 6.5, height = 6.5,
  plot = ggplot(od, aes(x = az)) +
    geom_histogram(breaks = seq(0, 360, 22.5), fill = "grey35", color = "white") +
    coord_polar(start = 0) + scale_x_continuous(limits = c(0,360),
      breaks = seq(0, 315, 45), labels = c("N","NE","E","SE","S","SW","W","NW")) +
    labs(title = "Orientacio de facanes (azimut fotogrametric)",
         subtitle = sprintf("n = %d | mitjana circular = %.0f | R = %.2f",
                            length(az), mitja, Rbar), x = NULL, y = "n") +
    theme_minimal(base_size = 11))

sink()
cat("FET. Figures, taules i resultats_analisi.txt a output/.\n")

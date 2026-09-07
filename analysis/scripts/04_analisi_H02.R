# ============================================================
# 04_analisi_H02.R - Contingencia suport geologic x construccio (H02)
# TFM Diablo Wasi (BD v26). Executar des d'analysis/:
#   Rscript scripts/04_analisi_H02.R
# Requereix: data/ generat per 00+01 i paquets ggplot2.
# Poblacio i clusters: identics a 02 (D1-D4, sessio 2026-09-07).
# ============================================================
suppressPackageStartupMessages(library(ggplot2))
set.seed(26)
for (d in c("output/figures","output/logs","output/tables"))
  dir.create(d, recursive = TRUE, showWarnings = FALSE)

M  <- read.csv("data/matriu_AX_cinc_valors.csv", check.names = FALSE)
TS <- read.csv("data/T_STRUCTURES.csv")
SU <- read.csv("data/L_SUPPORT.csv")

ll <- c("B","C","D","E","F","G","I","J","K","L","M","N","O","Q","R","S","T","V","X","Z")
pobl <- M$Record_Class %in% c("Built funerary structure","Natural funerary context") &
        rowSums(!is.na(M[, ll])) >= 10
A <- M[pobl, ]
X <- as.matrix(A[, ll])
P <- ifelse(is.na(X), NA, ifelse(X %in% c(1,2,3), 1, ifelse(X == 0, 0, NA)))
P <- matrix(as.numeric(P), nrow = nrow(X), dimnames = list(A$Code, ll))
jac <- function(P) { n <- nrow(P)
  D <- matrix(NA_real_, n, n, dimnames = list(rownames(P), rownames(P)))
  for (i in 1:n) for (j in i:n) {
    ok <- !is.na(P[i,]) & !is.na(P[j,])
    a <- sum(P[i,ok]==1 & P[j,ok]==1); b <- sum(P[i,ok]!=P[j,ok]); u <- a+b
    D[i,j] <- D[j,i] <- if (u==0) 0 else 1 - a/u }
  as.dist(D) }
A$Cluster <- factor(cutree(hclust(jac(P), method = "average"), 5)[A$Code])
A$Suport  <- SU$Name[match(TS$ID_Support[match(A$Code, TS$Code)], SU$ID)]

sink("output/logs/resultats_H02.txt", split = TRUE)
cat("=== H02: Suport geologic x construccio ===\n\nSuport x Tipologia:\n")
t1 <- table(A$Suport, A$Typology); print(t1)
c1 <- chisq.test(t1, simulate.p.value = TRUE, B = 20000)
cat("Chi2 MC: X2 =", round(c1$statistic, 1), " p =", signif(c1$p.value, 3), "\n\n")
cat("Suport x Cluster constructiu:\n")
t2 <- table(A$Suport, A$Cluster); print(t2)
c2 <- chisq.test(t2, simulate.p.value = TRUE, B = 20000)
cat("Chi2 MC: X2 =", round(c2$statistic, 1), " p =", signif(c2$p.value, 3), "\n")
sink()

d <- as.data.frame(t2); names(d) <- c("Suport","Cluster","n"); d <- d[d$n > 0, ]
ggsave("output/figures/fig_H02_suport_cluster.pdf", width = 8.5, height = 5,
  plot = ggplot(d, aes(Cluster, Suport, size = n, label = n)) +
    geom_point(shape = 21, fill = "grey70") +
    geom_text(size = 3, vjust = -1.6) +
    scale_size_area(max_size = 14, guide = "none") +
    labs(title = "Suport geologic x cluster constructiu (n = 92)",
         subtitle = sprintf("Chi2 Monte Carlo = %.1f, p = %s",
                            c2$statistic, format(signif(c2$p.value, 3))),
         x = "Cluster", y = NULL) +
    theme_minimal(base_size = 11))
cat("FET H02.\n")

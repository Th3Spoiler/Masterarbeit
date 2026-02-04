############################################################
# Masterarbeit
# Erstellung der Diagramme
#
# Autor: Maximilian Mäder
# Datum: 02/2026
#
# Hinweis:
# Dieses Skript dient der Erstellungen von Grafiken von der Berechnung für die Opportunitätsanalysen
# und ist Bestandteil des Anhangs der Masterarbeit.
# Eine Ausführung erfordert lokale Daten und Pfadanpassungen.
# Es wurde mit künstlicher Intelligenz (chat.ai der GWDG / Qen 3 30B A3B Instruct 2507) erstellt und von mir persönlich angepasst. 
############################################################

library(readxl)
library(dplyr)
library(ggplot2)
library(grid)
library(tidyr)

# ========================
# Plot-Nummern
# ========================
plot_ids <- list(
  hoehe            = "01",
  bhd              = "02",
  opp_box          = "03",
  opp_box_mensch   = "04",
  opp_box_computer = "05",
  avg_rel          = "06",
  cummean_opp      = "07"
)

# ========================
# Zielordner & Diskontfaktoren
# ========================
base_plot_dir <- "C:/Users/Max/Documents/R_Masterarbeit/Masterarbeit/Plots"

disk_rates <- c(
  "disk00" = 0.000,
  "disk05" = 0.005,
  "disk10" = 0.010,
  "disk15" = 0.015
)

zins_label <- function(rate_name) {
  zins <- as.numeric(sub("disk", "", rate_name)) / 10
  paste0("Zins = ", zins, " % p. A.")
}

for (rate_name in names(disk_rates)) {
  dir_path <- file.path(base_plot_dir, rate_name)
  if (!dir.exists(dir_path)) dir.create(dir_path, recursive = TRUE)
}

# ========================
# Farben & Theme
# ========================
custom_colors <- c("Mensch" = "green4", "Computer" = "black")

custom_theme <- theme_minimal(base_family = "Arial") +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 11),
    axis.text.y = element_text(size = 11),
    axis.title  = element_text(size = 11),
    
    plot.title = element_text(
      size = 11,
      margin = margin(b = 12)
    ),
    
    legend.position = "bottom",
    legend.title = element_blank(),
    legend.text  = element_text(size = 11),
    
    axis.line = element_line(color = "black"),
    axis.ticks = element_line(color = "black"),
    axis.ticks.length = unit(3, "mm"),
    panel.grid.minor = element_blank()
  )

# ========================
# Plot-Funktionen
# ========================
make_lineplot <- function(data, title, ylabel, filename, rate_name,
                          show_all_years = FALSE) {
  if (nrow(data) == 0) return(NULL)
  
  p <- ggplot(data, aes(AusscheideJahr, value, color = Player, group = Player)) +
    geom_line(linewidth = 0.8) +
    geom_point(size = 2) +
    geom_hline(yintercept = 0, color = "red", linetype = "dashed") +
    scale_color_manual(values = custom_colors, name = NULL, drop = FALSE) +
    labs(title = title, x = "Ausscheidejahr", y = ylabel) +
    custom_theme
  
  if (show_all_years) {
    p <- p + scale_x_continuous(breaks = sort(unique(data$AusscheideJahr)))
  }
  
  ggsave(
    file.path(base_plot_dir, rate_name, paste0(filename, ".png")),
    p, width = 8, height = 5, dpi = 300
  )
}

make_boxplot <- function(data, title, ylabel, filename, rate_name) {
  if (nrow(data) == 0) return(NULL)
  
  p <- ggplot(data, aes(factor(AusscheideJahr), value, fill = Player)) +
    geom_boxplot(outlier.shape = NA, alpha = 0.5,
                 position = position_dodge(width = 0.75)) +
    geom_jitter(aes(color = Player),
                position = position_jitterdodge(0.2, 0.75),
                alpha = 0.4, size = 1.2) +
    stat_summary(fun = mean, geom = "point",
                 shape = 21, size = 2.8,
                 fill = "white", color = "black",
                 position = position_dodge(width = 0.75)) +
    geom_hline(yintercept = 0, color = "red", linetype = "dashed") +
    scale_fill_manual(values = custom_colors, name = NULL, drop = FALSE) +
    scale_color_manual(values = custom_colors, guide = "none", drop = FALSE) +
    labs(title = title, x = "Ausscheidejahr", y = ylabel) +
    custom_theme
  
  ggsave(
    file.path(base_plot_dir, rate_name, paste0(filename, ".png")),
    p, width = 10, height = 5, dpi = 300
  )
}

make_boxplot_single <- function(data, title, ylabel, filename, rate_name,
                                fill_color = "grey70",
                                point_color = "black") {
  if (nrow(data) == 0) return(NULL)
  
  p <- ggplot(data, aes(factor(AusscheideJahr), value)) +
    geom_boxplot(outlier.shape = NA, fill = fill_color, alpha = 0.6) +
    geom_jitter(width = 0.2, alpha = 0.4, size = 1.2, color = point_color) +
    stat_summary(aes(group = Player),
      fun = mean, geom = "point",
                 shape = 21, size = 2.8,
                 fill = "white", color = "black") +
    geom_hline(yintercept = 0, color = "red", linetype = "dashed") +
    labs(title = title, x = "Ausscheidejahr", y = ylabel) +
    custom_theme
  
  ggsave(
    file.path(base_plot_dir, rate_name, paste0(filename, ".png")),
    p, width = 10, height = 5, dpi = 300
  )
}

# ========================
# Daten einlesen
# ========================
data_path <- "C:/Users/Max/Documents/R_Masterarbeit/Masterarbeit/Opportunitätsanalysen"

excel_files <- list.files(data_path, "\\.xlsx$", full.names = TRUE)
excel_files <- excel_files[!grepl("~\\$", basename(excel_files))]

alle_daten <- lapply(excel_files, function(file) {
  lapply(excel_sheets(file), function(sh) {
    read_excel(file, sheet = sh) %>% mutate(Sheet = sh)
  }) %>% bind_rows()
}) %>% bind_rows()

# Player-Namen robust vereinheitlichen
alle_daten <- alle_daten %>%
  mutate(
    Player = trimws(Player),
    Player = case_when(
      tolower(Player) == "mensch"   ~ "Mensch",
      tolower(Player) == "computer" ~ "Computer",
      TRUE ~ Player
    )
  )

alle_daten <- alle_daten %>%
  mutate(
    EntnahmeKategorie = case_when(
      AusscheideGrund == "-1" ~ "Keine_Entnahme",
      AusscheideGrund == "MortalitÃ¤t" ~ "Mortalität",
      AusscheideGrund %in% c("Ernte", "Durchforstung") ~ "Entnahme",
      TRUE ~ "Sonstige"
    )
  )

# ========================
# Auswertung
# ========================
for (bestand in c("A", "B")) {
  for (rate_name in names(disk_rates)) {
    
    diff_col_5    <- paste0("Wert_diff_", rate_name, "_5")
    diff_col_5pct <- paste0("Wert_diff_", rate_name, "_5_pct")
    
    if (!(diff_col_5 %in% names(alle_daten))) next
    
    basis <- alle_daten %>%
      filter(
        Bestand == bestand,
        Sheet == "Gesamt_Tote",
        EntnahmeKategorie == "Entnahme",
        Wert_now > 0
      ) %>%
      mutate(
        Wertzuwachs_rel = get(diff_col_5pct) * 100
      )
    
    # 01 Höhe (ohne Zinsangabe)
    make_boxplot(
      basis %>% mutate(value = Hoehe_m),
      paste0("Höhe im Ausscheidejahr – Bestand ", bestand),
      "Höhe (m)",
      paste0(plot_ids$hoehe, "_Boxplot_Hoehe_Bestand",
             bestand, "_", rate_name),
      rate_name
    )
    
    # 02 BHD (ohne Zinsangabe)
    make_boxplot(
      basis %>% mutate(value = BHD_mR_cm),
      paste0("BHD im Ausscheidejahr – Bestand ", bestand),
      "BHD (cm)",
      paste0(plot_ids$bhd, "_Boxplot_BHD_Bestand",
             bestand, "_", rate_name),
      rate_name
    )
    
    # 03 gemeinsamer jährlicher Opportunitätswert
    plot_box <- basis %>% mutate(value = get(diff_col_5))
    make_boxplot(
      plot_box,
      paste0("Jährlicher Opportunitätswert – Bestand ", bestand,
             " (", zins_label(rate_name), ")"),
      "€ pro Jahr",
      paste0(plot_ids$opp_box, "_Boxplot_Opportunitaetswert_Bestand",
             bestand, "_", rate_name),
      rate_name
    )
    
    # 04 Mensch (grün)
    make_boxplot_single(
      plot_box %>% filter(Player == "Mensch"),
      paste0("Jährlicher Opportunitätswert – Mensch – Bestand ", bestand,
             " (", zins_label(rate_name), ")"),
      "€ pro Jahr",
      paste0(plot_ids$opp_box_mensch,
             "_Boxplot_Opportunitaetswert_Mensch_Bestand",
             bestand, "_", rate_name),
      rate_name,
      fill_color = "green4",
      point_color = "green4"
    )
    
    # 05 Computer
    make_boxplot_single(
      plot_box %>% filter(Player == "Computer"),
      paste0("Jährlicher Opportunitätswert – Computer – Bestand ", bestand,
             " (", zins_label(rate_name), ")"),
      "€ pro Jahr",
      paste0(plot_ids$opp_box_computer,
             "_Boxplot_Opportunitaetswert_Computer_Bestand",
             bestand, "_", rate_name),
      rate_name
    )
    
    # 06 Ø jährlicher relativer Wertzuwachs
    make_lineplot(
      basis %>%
        group_by(AusscheideJahr, Player) %>%
        summarise(value = mean(Wertzuwachs_rel, na.rm = TRUE),
                  .groups = "drop"),
      paste0("Ø jährliche relative Wertzuwachsopportunität – Bestand ",
             bestand, " (", zins_label(rate_name), ")"),
      "% pro Jahr",
      paste0(plot_ids$avg_rel,
             "_Avg_rel_Wertzuwachs_Bestand",
             bestand, "_", rate_name),
      rate_name,
      show_all_years = TRUE
    )
    
    # 07 Kumulativer Ø jährlicher Opportunitätswert
    make_lineplot(
      basis %>%
        arrange(Player, AusscheideJahr) %>%
        group_by(Player) %>%
        mutate(value = cumsum(get(diff_col_5)) / row_number()) %>%
        ungroup() %>%
        group_by(AusscheideJahr, Player) %>%
        summarise(value = mean(value, na.rm = TRUE),
                  .groups = "drop"),
      paste0("Kumulativer Ø jährlicher Opportunitätswert der Ausscheidebäume – Bestand ",
             bestand, " (", zins_label(rate_name), ")"),
      "€ pro Baum und Jahr",
      paste0(plot_ids$cummean_opp,
             "_Kumulativ_Opportunitaetswert_Bestand",
             bestand, "_", rate_name),
      rate_name,
      show_all_years = TRUE
    )
  }
}

cat("Alle Diagramme wurden erfolgreich erstellt.\n")
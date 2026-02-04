############################################################
# Masterarbeit
# Opportunitätsanalyse für Waldbestände (WaldWeiser)
#
# Autor: Maximilian Mäder
# Datum: 02/2026
#
# Hinweis:
# Dieses Skript dient der Dokumentation der Berechnungen
# und ist Bestandteil des Anhangs der Masterarbeit.
# Eine Ausführung erfordert lokale Daten und Pfadanpassungen.
# Es wurde mit künstlicher Intelligenz (chat.ai der GWDG / Qen 3 30B A3B Instruct 2507) erstellt und von mir persönlich angepasst. 
############################################################

# ==========================================================
# 1. Pakete
# ==========================================================

library(xml2)
library(dplyr)
library(purrr)
library(readxl)
library(writexl)
library(stringr)
library(tools)

# ==========================================================
# 2. Hilfsfunktionen
# ==========================================================

## ---- 1: Formeln aus Settings-Datei laden (HandledLikeCode berücksichtigt) ----
growth_formulas_lookup_all <- function(settings_file) {
  doc <- read_xml(settings_file)
  species_nodes <- xml_find_all(doc, "//SpeciesDefinition")
  
  # Holt Formel, läuft durch HandledLikeCode-Kette
  resolve_formula <- function(code, tag) {
    visited <- c()
    current <- code
    repeat {
      if (current %in% visited) break
      visited <- c(visited, current)
      val <- xml_text(xml_find_first(doc,
                                     paste0("//SpeciesDefinition[Code='", current, "']/", tag)))
      if (!is.na(val) && val != "") return(val)
      handled <- xml_text(xml_find_first(doc,
                                         paste0("//SpeciesDefinition[Code='", current, "']/HandledLikeCode")))
      if (handled == "" || is.na(handled)) break
      current <- handled
    }
    return(NA_character_)
  }
  
  map_dfr(species_nodes, function(sp) {
    code <- as.integer(xml_text(xml_find_first(sp, "./Code")))
    tibble(
      species_code = code,
      vol_formula        = resolve_formula(code, "VolumeFunctionXML"),
      height_inc_formula = resolve_formula(code, "HeightIncrement"),
      diam_inc_formula   = resolve_formula(code, "DiameterIncrement"),
      ihpot_formula      = resolve_formula(code, "PotentialHeightIncrement")
    )
  })
}

## ---- 1b: Manueller Fallback ----
manual_fallback <- function(formulas_all, target_code, source_code) {
  source_form <- formulas_all %>% filter(species_code == source_code)
  idx <- which(formulas_all$species_code == target_code)
  for (col in names(formulas_all)[-1]) {
    formulas_all[[col]][idx] <- source_form[[col]]
  }
  formulas_all
}

## ---- 2: Sichere Formelauswertung ----
safe_eval_num <- function(formula_txt, env, label="") {
  if (is.null(formula_txt) || formula_txt == "" || all(is.na(formula_txt))) return(NA_real_)
  f_clean <- gsub("/\\*.*?\\*/", "", formula_txt)   # Kommentare raus
  f_clean <- gsub("if(", "ifelse(", f_clean, fixed = TRUE)
  f_clean <- gsub("ln", "log", f_clean)
  val <- try(eval(parse(text = f_clean), envir = env), silent = TRUE)
  if (inherits(val, "try-error") || is.null(val) || !is.numeric(val)) return(NA_real_)
  as.numeric(val)
}

# ==========================================================
# 3. Datenimport
# ==========================================================

## ---- 3: Inventur vollständig laden ----
parse_inventory_xml_full <- function(file) {
  doc <- read_xml(file)
  baum_nodes <- xml_find_all(doc, "//Baum")
  
  if (length(baum_nodes) == 0) return(tibble())
  
  baum_list <- lapply(baum_nodes, function(b) {
    kids <- xml_children(b)
    vals <- xml_text(kids)
    names(vals) <- xml_name(kids)
    as.list(vals)
  })
  
  df <- bind_rows(baum_list) %>%
    mutate(across(everything(), ~na_if(.x, ""))) %>%
    mutate(
      BaumartcodeLokal = as.integer(BaumartcodeLokal),
      BHD_mR_cm = as.numeric(BHD_mR_cm),
      Hoehe_m = as.numeric(Hoehe_m),
      Alter_Jahr = as.integer(Alter_Jahr),
      Kronenansatz_m = as.numeric(Kronenansatz_m),
      MittlererKronenDurchmesser_m = as.numeric(MittlererKronenDurchmesser_m),
      SiteIndex_m = as.numeric(SiteIndex_m),
      AusscheideJahr = as.integer(AusscheideJahr),
      Volumen_cbm = as.numeric(Volumen_cbm)
    ) %>%
    # Neuer eindeutiger Schlüssel: Baumartcode + Kennung + AusscheideJahr
    mutate(
      BaumKey = paste0(BaumartcodeLokal, "_", Kennung, "_", AusscheideJahr)
    )
  
  return(df)
}

## ---- 4: Species-Lookup berechnen ----
species_lookup_from_inv <- function(all_inv) {
  all_inv %>%
    group_by(BaumartcodeLokal) %>%
    summarise(
      sp.dg = sqrt(mean(BHD_mR_cm^2, na.rm=TRUE)),
      sp.hg = mean(Hoehe_m[which.min(abs(BHD_mR_cm -
                                           sqrt(mean(BHD_mR_cm^2, na.rm=TRUE))))], na.rm=TRUE),
      sp.h100 = mean(sort(Hoehe_m, decreasing=TRUE)[1:min(100, n())], na.rm=TRUE),
      sp.BHD_STD = sd(BHD_mR_cm, na.rm=TRUE),
      sp.Cw_dg = mean(MittlererKronenDurchmesser_m[which.min(abs(BHD_mR_cm -
                                                                   sqrt(mean(BHD_mR_cm^2, na.rm=TRUE))))], na.rm=TRUE),
      sp.gha = sum(pi * (BHD_mR_cm / 200)^2, na.rm=TRUE),
      .groups = "drop"
    ) %>%
    mutate(st.gha = sum(sp.gha, na.rm=TRUE))
}

## ---- 5: Preis-Lookup ----
price_file <- "C:/Users/Max/Documents/R_Masterarbeit/Masterarbeit/woodValuationDE_net_revenues_Unity.xlsx"
price_table <- read_excel(price_file) %>%
  mutate(
    species_code = as.integer(species_code),
    diameter_q_cm = as.numeric(diameter_q_cm),
    net_revenue_eur_per_m3 = as.numeric(net_revenue_eur_per_m3)
  )

lookup_price <- function(code, diameter) {
  subset <- filter(price_table, species_code == code)
  if (nrow(subset) == 0 || is.na(diameter)) return(NA_real_)
  subset$net_revenue_eur_per_m3[which.min(abs(subset$diameter_q_cm - diameter))]
}

## ---- 6: Umgebung für einen Baum erstellen ----
make_env_for_tree <- function(row, sp_lookup, form=NULL) {
  env <- new.env()
  env$t.d  <- row$BHD_mR_cm
  env$t.h  <- row$Hoehe_m
  env$t.age <- row$Alter_Jahr
  env$t.si <- row$SiteIndex_m
  env$t.cw <- row$MittlererKronenDurchmesser_m
  env$t.cb <- row$Kronenansatz_m
  env$t.c66 <- 0
  env$t.c66xy <- 0
  env$t.c66cxy <- 0
  env$random <- runif(1)
  
  # sp.* Variablen laden
  sp <- sp_lookup %>% filter(BaumartcodeLokal == row$BaumartcodeLokal)
  if (nrow(sp) > 0) {
    env$sp.dg     <- sp$sp.dg
    env$sp.hg     <- sp$sp.hg
    env$sp.h100   <- sp$sp.h100
    env$sp.BHD_STD <- sp$sp.BHD_STD
    env$sp.Cw_dg  <- sp$sp.Cw_dg
    env$sp.gha    <- sp$sp.gha
    env$st.gha    <- sp$st.gha
  }
  
  # t.hinc aus ihpot_formula berechnen
  if (!is.null(form) && !is.na(form$ihpot_formula) && form$ihpot_formula != "") {
    env$t.ihpot <- safe_eval_num(form$ihpot_formula, env)
    env$t.hinc <- env$t.ihpot
  } else {
    env$t.hinc <- NA_real_
  }
  
  return(env)
}

# ==========================================================
# 4. Berechnung
# ==========================================================

## ---- 7: Zukunftswert eines Baumes berechnen (mit 4 Diskontfaktoren: 0%, 0.5%, 1.0%, 1.5%) ----
calculate_future_value <- function(tree_row, formulas, sp_lookup,
                                   step_years = 5,
                                   discount_rates = c(0.000, 0.005, 0.010, 0.015)) {
  
  form <- formulas %>% filter(species_code == tree_row$BaumartcodeLokal)
  if (nrow(form) == 0) return(NULL)
  
  env_now <- make_env_for_tree(tree_row, sp_lookup, form)
  
  # Höhe und Durchmesser wachsen
  h_inc <- safe_eval_num(form$height_inc_formula, env_now)
  h_inc <- ifelse(is.na(h_inc), 0, h_inc)
  
  d_inc_raw <- safe_eval_num(form$diam_inc_formula, env_now)
  d_inc <- ifelse(is.na(d_inc_raw), 0, d_inc_raw * 100)  # Meter → Zentimeter
  
  # Umgebung für Zukunft
  env_future <- make_env_for_tree(tree_row, sp_lookup, form)
  env_future$t.d <- tree_row$BHD_mR_cm + d_inc
  env_future$t.h <- tree_row$Hoehe_m + h_inc
  
  # Volumen jetzt und in Zukunft
  vol_now <- safe_eval_num(form$vol_formula, env_now)
  vol_future <- safe_eval_num(form$vol_formula, env_future)
  
  # Preis jetzt und in Zukunft
  preis_now <- lookup_price(tree_row$BaumartcodeLokal, env_now$t.d)
  preis_future <- lookup_price(tree_row$BaumartcodeLokal, env_future$t.d)
  
  # Wert jetzt
  wert_now <- vol_now * preis_now
  
  # Wert in Zukunft mit Diskontierung
  wert_future_disk00 <- vol_future * preis_future  # 0 % → keine Diskontierung
  wert_future_disk05 <- vol_future * preis_future / ((1 + discount_rates[2]) ^ step_years)
  wert_future_disk10 <- vol_future * preis_future / ((1 + discount_rates[3]) ^ step_years)
  wert_future_disk15 <- vol_future * preis_future / ((1 + discount_rates[4]) ^ step_years)
  
  # Differenzen
  wert_diff_disk00 <- wert_future_disk00 - wert_now
  wert_diff_disk05 <- wert_future_disk05 - wert_now
  wert_diff_disk10 <- wert_future_disk10 - wert_now
  wert_diff_disk15 <- wert_future_disk15 - wert_now
  
  # Neue SPALTEN: Jährliche Durchschnittsrendite über 5 Jahre
  # 1. Wert_diff / 5
  wert_diff_disk00_5 <- wert_diff_disk00 / 5
  wert_diff_disk05_5 <- wert_diff_disk05 / 5
  wert_diff_disk10_5 <- wert_diff_disk10 / 5
  wert_diff_disk15_5 <- wert_diff_disk15 / 5
  
  # 2. Prozentwert: (jährliche Differenz / aktueller Wert)
  wert_diff_disk00_5_pct <- ifelse(wert_now == 0, NA_real_, wert_diff_disk00_5 / wert_now)
  wert_diff_disk05_5_pct <- ifelse(wert_now == 0, NA_real_, wert_diff_disk05_5 / wert_now)
  wert_diff_disk10_5_pct <- ifelse(wert_now == 0, NA_real_, wert_diff_disk10_5 / wert_now)
  wert_diff_disk15_5_pct <- ifelse(wert_now == 0, NA_real_, wert_diff_disk15_5 / wert_now)
  
  # Rückgabe: alle Spalten inkl. neuer
  tibble(
    BaumKey = tree_row$BaumKey,
    species_code = tree_row$BaumartcodeLokal,
    BHD_now = env_now$t.d,
    BHD_future = env_future$t.d,
    Vol_now = vol_now,
    Vol_future = vol_future,
    Wert_now = wert_now,
    
    # Explizite Diskontspalten (keine Doppelung!)
    Wert_future_disk00 = wert_future_disk00,
    Wert_future_disk05 = wert_future_disk05,
    Wert_future_disk10 = wert_future_disk10,
    Wert_future_disk15 = wert_future_disk15,
    
    Wert_diff_disk00 = wert_diff_disk00,
    Wert_diff_disk05 = wert_diff_disk05,
    Wert_diff_disk10 = wert_diff_disk10,
    Wert_diff_disk15 = wert_diff_disk15,
    
    # NEUE SPALTEN:
    Wert_diff_disk00_5 = wert_diff_disk00_5,
    Wert_diff_disk05_5 = wert_diff_disk05_5,
    Wert_diff_disk10_5 = wert_diff_disk10_5,
    Wert_diff_disk15_5 = wert_diff_disk15_5,
    
    Wert_diff_disk00_5_pct = wert_diff_disk00_5_pct,
    Wert_diff_disk05_5_pct = wert_diff_disk05_5_pct,
    Wert_diff_disk10_5_pct = wert_diff_disk10_5_pct,
    Wert_diff_disk15_5_pct = wert_diff_disk15_5_pct
  )
}

## ---- 8: Automatisches Einlesen aller XML-Dateien und Zusammenfassung in zwei Blättern ----

# Pfad zum Simulation-Ordner
sim_dir <- "C:/Users/Max/Documents/meineWaelder/Simulation"

# Liste alle .xml-Dateien
xml_files <- dir(sim_dir, pattern = "\\.xml$", full.names = TRUE)

if (length(xml_files) == 0) {
  stop("Keine .xml-Dateien im Ordner gefunden: ", sim_dir)
}

# Lade die Formeln (einmalig)
settings_path <- "C:/Users/Max/Documents/R_Masterarbeit/Masterarbeit/ForestSimulatorNWGermany6.xml"
formulas_all <- growth_formulas_lookup_all(settings_path)
formulas_all <- manual_fallback(formulas_all, 441, 411)  # Fallback für 441

# Preis-Tabelle (einmalig laden)
price_file <- "C:/Users/Max/Documents/R_Masterarbeit/Masterarbeit/woodValuationDE_net_revenues_Unity.xlsx"
price_table <- read_excel(price_file) %>%
  mutate(
    species_code = as.integer(species_code),
    diameter_q_cm = as.numeric(diameter_q_cm),
    net_revenue_eur_per_m3 = as.numeric(net_revenue_eur_per_m3)
  )

# Listen für Ergebnisse
all_results_tote <- list()      # Tote Bäume
all_results_lebende <- list()    # Lebende Bäume

# ==========================================================
# 5. Hauptprozess
# ==========================================================

# Schleife über alle XML-Dateien
for (xml_file in xml_files) {
  # Extrahiere Bestand und Player aus dem Dateinamen
  filename <- basename(xml_file)
  base_name <- file_path_sans_ext(filename)
  
  match <- regmatches(base_name, regexec("Bestand([A-Z])_([A-Za-z0-9]+)", base_name))
  
  if (length(match) > 0 && length(match[[1]]) >= 3) {
    bestand <- match[[1]][2]  # z. B. "A"
    player <- match[[1]][3]   # z. B. "Computer"
  } else {
    bestand <- NA_character_
    player <- NA_character_
    warning("Unbekanntes Dateiformat: ", filename, ". Ignoriere.")
    next
  }
  
  # Lade Inventur
  # Lade Inventur
  inventory_full <- parse_inventory_xml_full(xml_file)
  
  if (nrow(inventory_full) == 0) {
    message("Keine Bäume in ", filename, " gefunden.")
    next
  }
  
  # Zähle Bäume unter 10 cm
  n_under_10 <- sum(inventory_full$BHD_mR_cm < 10, na.rm = TRUE)
  if (n_under_10 > 0) {
    message("⚠️ ", n_under_10, " Bäume mit BHD < 10 cm in ", filename, " entfernt.")
  }
  
  # FILTER: Nur Bäume mit BHD >= 10 cm behalten
  inventory_full <- inventory_full %>%
    filter(BHD_mR_cm >= 10)
  
  if (nrow(inventory_full) == 0) {
    message("Keine Bäume mit BHD >= 10 cm in ", filename, " gefunden.")
    next
  }
  
  # Berechne Species-Lookup
  sp_lookup <- species_lookup_from_inv(inventory_full)
  
  # Tote und lebende Bäume trennen
  dead_trees <- inventory_full %>% filter(AusscheideJahr != -1)
  living_trees <- inventory_full %>% filter(AusscheideJahr == -1)
  
  # Tote Bäume: Berechne Zukunftswerte
  if (nrow(dead_trees) > 0) {
    results_dead_calc <- map_dfr(seq_len(nrow(dead_trees)), ~ {
      calculate_future_value(dead_trees[.x, ], formulas_all, sp_lookup, step_years = 5)
    })
    
    results_dead_full <- left_join(dead_trees, results_dead_calc, by = "BaumKey") %>%
      mutate(Bestand = bestand, Player = player)
    
    all_results_tote[[length(all_results_tote) + 1]] <- results_dead_full
  }
  
  # Lebende Bäume: Nur mit Bestand/Player hinzufügen
  if (nrow(living_trees) > 0) {
    living_trees <- living_trees %>%
      mutate(Bestand = bestand, Player = player)
    
    all_results_lebende[[length(all_results_lebende) + 1]] <- living_trees
  }
}

# Zusammenfassen aller Ergebnisse
if (length(all_results_tote) == 0) {
  warning("Keine toten Bäume gefunden. Tote-Bäume-Tabelle ist leer.")
}
if (length(all_results_lebende) == 0) {
  warning("Keine lebenden Bäume gefunden. Lebende-Bäume-Tabelle ist leer.")
}

results_tote_gesamt <- if (length(all_results_tote) > 0) {
  bind_rows(all_results_tote) %>%
    arrange(Bestand, Player, BaumKey)
} else {
  tibble()  # Leere Tabelle
}

results_lebende_gesamt <- if (length(all_results_lebende) > 0) {
  bind_rows(all_results_lebende) %>%
    arrange(Bestand, Player, BaumKey)
} else {
  tibble()  # Leere Tabelle
}

# Erstelle Ausgabedatei
output_file <- "Opportunitätsanalysen/Opportunitaetsanalyse_NoDisko_Gesamt.xlsx"

# Schreibe beide Tabellen in eine Excel-Datei
write_xlsx(
  list(
    Gesamt_Tote = results_tote_gesamt,
    Gesamt_Lebende = results_lebende_gesamt
  ),
  path = output_file
)

# Ausgabe
cat("Alle XML-Dateien verarbeitet.\n")
cat("   → Gespeichert als: '", output_file, "'\n")
cat("   → Tote Bäume: ", nrow(results_tote_gesamt), " Zeilen\n")
cat("   → Lebende Bäume: ", nrow(results_lebende_gesamt), " Zeilen\n")


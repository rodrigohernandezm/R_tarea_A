# ============================================================================
# Curso: Software para Analisis e Interpretacion de la Informacion para la
# Investigacion Cientifica (EIC401)
# Examen: Muestreo aplicado en R - Base de datos de tuberculosis
#
# Este script reproduce exactamente los dos procedimientos de muestreo
# solicitados con set.seed(2026) y presenta el procedimiento y los resultados
# necesarios para responder las 30 preguntas.
# ============================================================================

library(readxl)
library(dplyr)

options(width = 120)
SI <- "S\u00ed"

# Funcion auxiliar para imprimir porcentajes con dos decimales.
porcentaje <- function(numerador, denominador) {
  100 * numerador / denominador
}

# ============================================================================
# EJERCICIO 1 - MUESTREO ALEATORIO SIMPLE
# ============================================================================

# Codigo de muestreo indicado en el enunciado (sin modificaciones).
datos <- read_excel("database_tb_limpia.xlsx", sheet = "datos", guess_max = 5000)
datos <- as.data.frame(datos)

set.seed(2026)
n1 <- 300
indices1 <- sample(1:nrow(datos), n1)
muestra1 <- datos[indices1, ]

cat("\n============================================================\n")
cat("EJERCICIO 1 - MUESTREO ALEATORIO SIMPLE\n")
cat("============================================================\n")

# 1. Tamano de muestra y porcentaje respecto de N = 3,104.
n_muestra1 <- nrow(muestra1)
pct_muestra1 <- porcentaje(n_muestra1, nrow(datos))
cat(sprintf("\n1. Observaciones: %d; porcentaje de la base: %.2f%%.\n",
            n_muestra1, pct_muestra1))

# 2. Frecuencias por sexo y porcentaje de hombres.
tabla_sexo1 <- table(muestra1$sexo)
n_masc1 <- unname(tabla_sexo1["Masculino"])
n_fem1 <- unname(tabla_sexo1["Femenino"])
pct_masc1 <- porcentaje(n_masc1, n_muestra1)
cat(sprintf("\n2. Masculino: %d; Femenino: %d; hombres: %.2f%%.\n",
            n_masc1, n_fem1, pct_masc1))
print(tabla_sexo1)

# 3. Media, mediana y faltantes de edad al diagnostico.
media_edad1 <- mean(muestra1$edad_al_diagnostico, na.rm = TRUE)
mediana_edad1 <- median(muestra1$edad_al_diagnostico, na.rm = TRUE)
na_edad1 <- sum(is.na(muestra1$edad_al_diagnostico))
cat(sprintf("\n3. Edad al diagnostico: media = %.2f; mediana = %.2f; NA = %d.\n",
            media_edad1, mediana_edad1, na_edad1))

# 4. Porcentaje de pacientes fallecidos segun pronostico.
n_fallecidos1 <- sum(muestra1$pronostico == "Fallecido", na.rm = TRUE)
pct_fallecidos1 <- porcentaje(n_fallecidos1, n_muestra1)
cat(sprintf("\n4. Fallecidos: %d (%.2f%% de la muestra).\n",
            n_fallecidos1, pct_fallecidos1))

# 5. Categoria modal de edad_rangos y numero de datos faltantes.
tabla_edad_rangos1 <- table(muestra1$edad_rangos)
edad_rango_modal1 <- names(which.max(tabla_edad_rangos1))
n_edad_rangos_disponibles1 <- sum(!is.na(muestra1$edad_rangos))
na_edad_rangos1 <- sum(is.na(muestra1$edad_rangos))
cat(sprintf(paste0("\n5. Categoria mas frecuente: %s (%d casos); ",
                   "NA: %d; datos disponibles: %d.\n"),
            edad_rango_modal1,
            unname(max(tabla_edad_rangos1)),
            na_edad_rangos1,
            n_edad_rangos_disponibles1))
print(table(muestra1$edad_rangos, useNA = "ifany"))

# 6. Porcentaje con coinfeccion.
n_coinfeccion1 <- sum(muestra1$coinfeccion == SI, na.rm = TRUE)
pct_coinfeccion1 <- porcentaje(n_coinfeccion1, n_muestra1)
cat(sprintf("\n6. Coinfeccion = Si: %d (%.2f%%).\n",
            n_coinfeccion1, pct_coinfeccion1))

# 7. Diagnostico de VIH positivo: frecuencia y porcentaje.
n_vih1 <- sum(muestra1$dx_vih == SI, na.rm = TRUE)
pct_vih1 <- porcentaje(n_vih1, n_muestra1)
cat(sprintf("\n7. dx_vih = Si: %d (%.2f%%).\n", n_vih1, pct_vih1))

# 8. Clasificacion por tipo de TB y categoria modal.
tabla_tipo_tb1 <- table(muestra1$tipo_de_tb, useNA = "ifany")
tipo_tb_modal1 <- names(which.max(tabla_tipo_tb1))
cat(sprintf("\n8. Categoria mas frecuente: %s (%d casos).\n",
            tipo_tb_modal1, unname(max(tabla_tipo_tb1))))
print(tabla_tipo_tb1)

# 9. Valores disponibles y media de cd4.
n_cd4_disponibles1 <- sum(!is.na(muestra1$cd4))
media_cd4_1 <- mean(muestra1$cd4, na.rm = TRUE)
cat(sprintf("\n9. cd4 disponibles: %d; media: %.2f celulas/ml.\n",
            n_cd4_disponibles1, media_cd4_1))

# 10. Ano con mayor numero de pacientes.
tabla_anio1 <- table(muestra1$anio)
anio_modal1 <- names(which.max(tabla_anio1))
cat(sprintf("\n10. Ano con mayor frecuencia: %s (%d pacientes).\n",
            anio_modal1, unname(max(tabla_anio1))))
print(tabla_anio1)

# 11. Casos Activos entre los pacientes con seguimiento registrado.
n_seguimiento_disponible1 <- sum(!is.na(muestra1$seguimiento))
n_activos1 <- sum(muestra1$seguimiento == "Activo", na.rm = TRUE)
cat(sprintf("\n11. Seguimiento disponible: %d; estado Activo: %d.\n",
            n_seguimiento_disponible1, n_activos1))

# 12. Baciloscopia positiva y datos faltantes.
n_bacilo_si1 <- sum(muestra1$baciloscopia == SI, na.rm = TRUE)
pct_bacilo_si1 <- porcentaje(n_bacilo_si1, n_muestra1)
na_bacilo1 <- sum(is.na(muestra1$baciloscopia))
cat(sprintf("\n12. Baciloscopia positiva: %d (%.2f%%); NA: %d.\n",
            n_bacilo_si1, pct_bacilo_si1, na_bacilo1))

# 13. Contacto con un caso de tuberculosis y porcentaje.
# Nota: el nombre de la variable en la base es 'conctacto', tal como aparece
# en el enunciado, aunque contiene un error ortografico.
n_contacto_si1 <- sum(muestra1$conctacto == SI, na.rm = TRUE)
pct_contacto_si1 <- porcentaje(n_contacto_si1, n_muestra1)
cat(sprintf("\n13. conctacto = Si: %d (%.2f%%).\n",
            n_contacto_si1, pct_contacto_si1))

# 14. Media y faltantes de meses_diagnostico_vih.
media_meses_vih1 <- mean(muestra1$meses_diagnostico_vih, na.rm = TRUE)
na_meses_vih1 <- sum(is.na(muestra1$meses_diagnostico_vih))
cat(sprintf("\n14. meses_diagnostico_vih: media = %.2f meses; NA = %d.\n",
            media_meses_vih1, na_meses_vih1))

# 15. Frecuencias absolutas y relativas de cultivo, incluyendo NA.
tabla_cultivo_abs1 <- table(muestra1$cultivo, useNA = "ifany")
tabla_cultivo_rel1 <- prop.table(tabla_cultivo_abs1)
tabla_cultivo1 <- data.frame(
  categoria = names(tabla_cultivo_abs1),
  frecuencia = as.integer(tabla_cultivo_abs1),
  proporcion = as.numeric(tabla_cultivo_rel1),
  porcentaje = 100 * as.numeric(tabla_cultivo_rel1),
  row.names = NULL
)
cat("\n15. Tabla de cultivo (incluye NA):\n")
print(tabla_cultivo1, digits = 4)

# ============================================================================
# EJERCICIO 2 - MUESTREO ESTRATIFICADO POR SEXO
# ============================================================================

# Codigo de muestreo indicado en el enunciado (sin modificaciones).
datos <- read_excel("database_tb_limpia.xlsx", sheet = "datos", guess_max = 5000)
datos <- as.data.frame(datos)

set.seed(2026)
n_total <- 300
prop_masc <- sum(datos$sexo == "Masculino") / nrow(datos)
n_masc <- round(n_total * prop_masc)
n_fem  <- n_total - n_masc

datos_masc <- datos[datos$sexo == "Masculino", ]
datos_fem  <- datos[datos$sexo == "Femenino", ]

muestra_masc <- datos_masc[sample(1:nrow(datos_masc), n_masc), ]
muestra_fem  <- datos_fem[sample(1:nrow(datos_fem), n_fem), ]

muestra2 <- rbind(muestra_masc, muestra_fem)

cat("\n============================================================\n")
cat("EJERCICIO 2 - MUESTREO ESTRATIFICADO POR SEXO\n")
cat("============================================================\n")

# Tabla auxiliar con denominadores por sexo.
n_por_sexo2 <- table(muestra2$sexo)

# 1. Distribucion estratificada y comparacion con la base completa.
tabla_sexo_base <- table(datos$sexo)
prop_sexo_base <- prop.table(tabla_sexo_base)
tabla_sexo2 <- table(muestra2$sexo)
prop_sexo2 <- prop.table(tabla_sexo2)
cat(sprintf(paste0("\n1. Muestra2: Masculino = %d (%.2f%%), Femenino = %d (%.2f%%). ",
                   "Base completa: Masculino = %.2f%%, Femenino = %.2f%%.\n"),
            tabla_sexo2["Masculino"], 100 * prop_sexo2["Masculino"],
            tabla_sexo2["Femenino"], 100 * prop_sexo2["Femenino"],
            100 * prop_sexo_base["Masculino"],
            100 * prop_sexo_base["Femenino"]))

# 2. Porcentaje de fallecidos dentro de cada sexo.
tabla_pronostico_sexo2 <- table(muestra2$sexo, muestra2$pronostico)
pct_fallecidos_sexo2 <- 100 * prop.table(tabla_pronostico_sexo2, margin = 1)[, "Fallecido"]
cat(sprintf("\n2. Fallecidos: mujeres = %.2f%%; hombres = %.2f%%.\n",
            pct_fallecidos_sexo2["Femenino"],
            pct_fallecidos_sexo2["Masculino"]))
print(tabla_pronostico_sexo2)

# 3. Sexo con mayor porcentaje de fallecidos.
sexo_mayor_fallecidos2 <- names(which.max(pct_fallecidos_sexo2))
cat(sprintf("\n3. Mayor porcentaje de fallecidos: %s (%.2f%%).\n",
            sexo_mayor_fallecidos2,
            max(pct_fallecidos_sexo2)))

# 4. Edad promedio al diagnostico por sexo.
edad_media_sexo2 <- aggregate(
  edad_al_diagnostico ~ sexo,
  data = muestra2,
  FUN = function(x) mean(x, na.rm = TRUE)
)
cat("\n4. Edad promedio al diagnostico por sexo:\n")
print(edad_media_sexo2)

# 5. Mediana de edad al diagnostico por sexo.
edad_mediana_sexo2 <- aggregate(
  edad_al_diagnostico ~ sexo,
  data = muestra2,
  FUN = function(x) median(x, na.rm = TRUE)
)
sexo_mayor_mediana2 <- edad_mediana_sexo2$sexo[which.max(edad_mediana_sexo2$edad_al_diagnostico)]
cat("\n5. Mediana de edad al diagnostico por sexo:\n")
print(edad_mediana_sexo2)
cat(sprintf("La mediana es mayor en: %s.\n", sexo_mayor_mediana2))

# 6. Porcentaje con diagnostico VIH positivo dentro de cada sexo.
tabla_vih_sexo2 <- table(muestra2$sexo, muestra2$dx_vih)
pct_vih_sexo2 <- 100 * prop.table(tabla_vih_sexo2, margin = 1)[, SI]
cat(sprintf("\n6. dx_vih = Si: hombres = %.2f%%; mujeres = %.2f%%.\n",
            pct_vih_sexo2["Masculino"], pct_vih_sexo2["Femenino"]))
print(tabla_vih_sexo2)

# 7. Tipo de TB mas frecuente por sexo y numero de casos.
tabla_tipo_tb_sexo2 <- table(muestra2$sexo, muestra2$tipo_de_tb)
tipo_tb_modal_sexo2 <- apply(tabla_tipo_tb_sexo2, 1, function(x) names(which.max(x)))
n_tipo_tb_modal_sexo2 <- apply(tabla_tipo_tb_sexo2, 1, max)
cat(sprintf(paste0("\n7. Tipo de TB mas frecuente: hombres = %s (%d casos); ",
                   "mujeres = %s (%d casos).\n"),
            tipo_tb_modal_sexo2["Masculino"], n_tipo_tb_modal_sexo2["Masculino"],
            tipo_tb_modal_sexo2["Femenino"], n_tipo_tb_modal_sexo2["Femenino"]))
print(tabla_tipo_tb_sexo2)

# 8. Porcentaje con coinfeccion dentro de cada sexo.
tabla_coinfeccion_sexo2 <- table(muestra2$sexo, muestra2$coinfeccion)
pct_coinfeccion_sexo2 <- 100 * prop.table(tabla_coinfeccion_sexo2, margin = 1)[, SI]
cat(sprintf("\n8. Coinfeccion = Si: mujeres = %.2f%%; hombres = %.2f%%.\n",
            pct_coinfeccion_sexo2["Femenino"],
            pct_coinfeccion_sexo2["Masculino"]))
print(tabla_coinfeccion_sexo2)

# 9. Casos con seguimiento Fallecido por sexo.
tabla_seguimiento_sexo2 <- table(muestra2$sexo, muestra2$seguimiento)
fallecido_seguimiento_sexo2 <- tabla_seguimiento_sexo2[, "Fallecido"]
sexo_mayor_seguimiento_fallecido2 <- names(which.max(fallecido_seguimiento_sexo2))
cat(sprintf(paste0("\n9. Seguimiento Fallecido: mujeres = %d; hombres = %d. ",
                   "Mayor numero: %s.\n"),
            fallecido_seguimiento_sexo2["Femenino"],
            fallecido_seguimiento_sexo2["Masculino"],
            sexo_mayor_seguimiento_fallecido2))

# 10. Valores NA de edad_rangos por sexo.
na_edad_rangos_sexo2 <- tapply(is.na(muestra2$edad_rangos), muestra2$sexo, sum)
cat(sprintf("\n10. NA en edad_rangos: mujeres = %d; hombres = %d.\n",
            na_edad_rangos_sexo2["Femenino"],
            na_edad_rangos_sexo2["Masculino"]))

# 11. Porcentaje con baciloscopia positiva dentro de cada sexo.
# El denominador es el total seleccionado de cada sexo; los NA no se cuentan
# como positivos. Tambien se reportan los NA para transparencia.
n_bacilo_si_sexo2 <- tapply(muestra2$baciloscopia == SI, muestra2$sexo,
                            function(x) sum(x, na.rm = TRUE))
na_bacilo_sexo2 <- tapply(is.na(muestra2$baciloscopia), muestra2$sexo, sum)
pct_bacilo_si_sexo2 <- 100 * n_bacilo_si_sexo2 / n_por_sexo2[names(n_bacilo_si_sexo2)]
cat(sprintf(paste0("\n11. Baciloscopia positiva: hombres = %.2f%% (%d NA); ",
                   "mujeres = %.2f%% (%d NA).\n"),
            pct_bacilo_si_sexo2["Masculino"], na_bacilo_sexo2["Masculino"],
            pct_bacilo_si_sexo2["Femenino"], na_bacilo_sexo2["Femenino"]))

# 12. Media de cd4 por sexo, solo con valores disponibles.
cd4_resumen_sexo2 <- muestra2 %>%
  group_by(sexo) %>%
  summarize(
    n_disponible = sum(!is.na(cd4)),
    media_cd4 = mean(cd4, na.rm = TRUE),
    .groups = "drop"
  )
sexo_mayor_cd4_2 <- cd4_resumen_sexo2$sexo[which.max(cd4_resumen_sexo2$media_cd4)]
cat("\n12. cd4 por sexo (valores disponibles):\n")
print(cd4_resumen_sexo2)
cat(sprintf(paste0("Media exacta: mujeres = %.2f (n = %d); ",
                   "hombres = %.2f (n = %d).\n"),
            cd4_resumen_sexo2$media_cd4[cd4_resumen_sexo2$sexo == "Femenino"],
            cd4_resumen_sexo2$n_disponible[cd4_resumen_sexo2$sexo == "Femenino"],
            cd4_resumen_sexo2$media_cd4[cd4_resumen_sexo2$sexo == "Masculino"],
            cd4_resumen_sexo2$n_disponible[cd4_resumen_sexo2$sexo == "Masculino"]))
cat(sprintf("La media de cd4 es mayor en: %s.\n", sexo_mayor_cd4_2))

# 13. Rango de edad mas frecuente dentro de cada sexo.
tabla_rango_edad_sexo2 <- table(muestra2$sexo, muestra2$edad_rangos)
rango_modal_sexo2 <- apply(tabla_rango_edad_sexo2, 1, function(x) names(which.max(x)))
n_rango_modal_sexo2 <- apply(tabla_rango_edad_sexo2, 1, max)
cat(sprintf(paste0("\n13. Rango de edad modal: mujeres = %s (%d casos); ",
                   "hombres = %s (%d casos).\n"),
            rango_modal_sexo2["Femenino"], n_rango_modal_sexo2["Femenino"],
            rango_modal_sexo2["Masculino"], n_rango_modal_sexo2["Masculino"]))
print(tabla_rango_edad_sexo2)

# 14. Casos con seguimiento Abandono por sexo.
abandono_sexo2 <- tabla_seguimiento_sexo2[, "Abandono"]
cat(sprintf("\n14. Seguimiento Abandono: mujeres = %d; hombres = %d.\n",
            abandono_sexo2["Femenino"], abandono_sexo2["Masculino"]))

# 15. La interpretacion narrativa se construye con los resultados anteriores.
cat(paste0(
  "\n15. Sintesis automatica: En esta muestra, el porcentaje de fallecidos fue ",
  sprintf("%.2f%% en mujeres y %.2f%% en hombres. ",
          pct_fallecidos_sexo2["Femenino"], pct_fallecidos_sexo2["Masculino"]),
  "Los casos de seguimiento Fallecido fueron ",
  fallecido_seguimiento_sexo2["Femenino"], " en mujeres y ",
  fallecido_seguimiento_sexo2["Masculino"], " en hombres, mientras que los abandonos fueron ",
  abandono_sexo2["Femenino"], " y ", abandono_sexo2["Masculino"],
  ", respectivamente. Estas diferencias son descriptivas de la muestra y no demuestran causalidad.\n"
))

# Controles finales de reproducibilidad y consistencia.
stopifnot(
  nrow(datos) == 3104,
  nrow(muestra1) == 300,
  nrow(muestra2) == 300,
  sum(tabla_sexo2) == 300,
  n_masc + n_fem == 300,
  all(indices1 >= 1 & indices1 <= nrow(datos))
)

cat("\nControles finales superados: N = 3,104 y ambas muestras tienen n = 300.\n")

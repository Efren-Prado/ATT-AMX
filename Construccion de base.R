library(urca)
library(readxl)
library(dplyr)
library(forecast)
library(fable)
library(fabletools)
library(tseries)
library(NonlinearTSA)
library(openxlsx)
library(tempdisagg)
library(tsbox)
library(tidyverse)
library(lubridate)
library(quantmod)
library(imputeTS)


# AMX:
# 
# Obtencion del precio de acciones de AMX:



# America Movil precio USD/MX (AMX) diaria/mensual ------------------------

# auto.assign = F me permite convertir los datos en un objeto y manejarlos
AMX.xts.d <- getSymbols("AMX", from = "2001-02-12", periodicity = "daily", auto.assign = F)

AMX.xts.m <- getSymbols("AMX", from = "2001-02-12", periodicity = "monthly", auto.assign = F)

#When auto.assign is set to TRUE (which is the default), the function automatically assigns the downloaded data to an object with the same name as the symbol being downloaded.

# Convertir a DF:
AMX.d.d <- data.frame(Date = index(AMX.xts.d), Close = AMX.xts.d$AMX.Close)
AMX.d.m <- data.frame(Date = index(AMX.xts.m), Close = AMX.xts.m$AMX.Close)

# AMX.d.d : Precio de acciones - AMX -Diario
# AMX.d.m: Precio de acciones - AMX - Mensual


# America Movil precio USD/MX (AMX) anual ---------------------------------

# Creo la secuecnia de años de 1901 a 2005:
years <- 2001:2023    # Creo el vector de años

AMX.d.a <- data.frame(years)  # Creo un df con ese vector

AMX.d.a <- AMX.d.a %>% mutate(precio = "") # Columna vacia

# Agregar los datos de la columna AMX.Close ordenandolos por año de la columna Date de la base AMX.d.m
AMX.d.a <- aggregate(AMX.Close ~ year(Date), data = AMX.d.d, mean)

# AMX.d.a: Precio de acciones - AMX - Anual


# Numero de acciones - mensual AMX ----------------------------------------

AMX.a.a <- read_xlsx("datos/acciones_corrientes_AMX_promedio_ponderado anual.xlsx")

# Borrar las ultimas dos filas: 
AMX.a.a <- slice(AMX.a.a, 1:(n() - 2))

AMX.a.a.ts <- ts(AMX.a.a$`Promedio ponderado de acciones en circulacion(millones)`, start = c(2001), frequency = 1)

# Desagregacion a mensual
ts_monthly_litterman <- td(AMX.a.a.ts ~ 1, method = "litterman-minrss", conversion = "mean", to = "month")

fechas <- seq(as.Date("2001-01-01"), as.Date("2021-12-31"), by = "month")

AMX.a.m.litterman <- data.frame(Date = fechas, `Acciones mensuales` = ts_monthly_litterman$values)

# AMX.a.a: Numero de acciones - AMX - Anual

# AMX.a.m.litterman: Numero de acciones - AMX - Mensual


# Numero de acciones - diario AMX -----------------------------------------

AMX.a.a <- read_xlsx("datos/acciones_corrientes_AMX_promedio_ponderado anual.xlsx")

# Borrar las ultimas dos filas: 
AMX.a.a <- slice(AMX.a.a, 1:(n() - 2))

AMX.a.a$Fecha <- as.Date(paste0(AMX.a.a$Fecha, "-01-01"), format = "%Y-%m-%d")

AMX.a.a.XTS <- xts(AMX.a.a$`Promedio ponderado de acciones en circulacion(millones)`, order.by = AMX.a.a$Fecha, frequency = 1)

# AMX.a.a.ts <- ts(AMX.a.a$`Promedio ponderado de acciones en circulacion(millones)`, start = c(2001), frequency = 1)



# PRECIO DE ACCIONES AT&T: -----------------------------------------------------

# PRECIO DE ACCIONES ANUAL-MENSUAL-DIARIO

# Cargo las bases
a <- read_excel("D:/Efren_Prado/Documents/Tesis/datos/Historicalprice_1901-1920.xlsx", range = "A8:C4095")
a1 <- read_excel("D:/Efren_Prado/Documents/Tesis/datos/Historicalprice_1920-1940.xlsx", range = "A8:C6274")
a2 <- read_excel("D:/Efren_Prado/Documents/Tesis/datos/Historicalprice_1940-1960.xlsx", range = "A8:C5848")
a3 <- read_excel("D:/Efren_Prado/Documents/Tesis/datos/Historicalprice_1960-1980.xlsx", range = "A8:C5351")
a4 <- read_excel("D:/Efren_Prado/Documents/Tesis/datos/Historicalprice_1980-2005.xlsx", range = "A8:C6654")

# Las uno y elimino duplicados:
base_chida <- rbind(a4,a3,a2,a1,a)
base2 <- base_chida[!duplicated(base_chida[,c("Date")]),]

# Base con solo high y date
base2 <- base2[, -2]

# Mensualizar los datos ---------------------------------------------------

# Cambia el formato de fecha
base2$Date <- as.Date(base2$Date, format="%b %d, %Y")  # Base de solo high con formato de fecha requerido

# Invierto el orden:
base2 <- base2[rev(1:nrow(base2)), ]
base2 <- base2[-1,]     # Borro la primera fila

base2   # Base de datos sin NAs pero con fechas faltantes   

# Solamente para ver la grafica
# > base2_xts <- as.xts(base2)
# > plot(base2_xts)

# Base completa  --------------------------------------------------------------

# Creo una base con TODAS LAS FECHAS
dates <- seq(as.Date("1901-09-20"), as.Date("2005-11-18"), by = "day")

# filter out weekends (Saturday and Sunday)
dates_filtered <- dates[!(weekdays(dates) %in% c("Saturday", "Sunday"))]

# create a dataframe with the filtered dates
df <- data.frame(Date = dates_filtered)

# merge with the df dataframe existente (base2)
merged_df <- merge(df, base2, by = "Date", all = TRUE)  # Base con datos diarios con NA's

# Eliminacion de ceros ----------------------------------------------------

# Remplazo los ceros con NA, para despues imputarlos
merged_df$High <- ifelse(merged_df$High == 0, NA, merged_df$High)

# Interpolacion con imputeTS ----------------------------------------------

# Imputación utilizando na_interpolation
serie_interpolada <- na_interpolation(merged_df$High, option = "spline")

# Imputación utilizando na_kalman
serie_kalman <- na_kalman(merged_df$High)

# KALMAN usa los siguientes valores predeterminados:
# na_kalman(x, model = "StructTS", smooth = TRUE, level = NULL, slope = NULL,
#           season = NULL, cycle = NULL, fill_method = "auto", na_only = FALSE)

# Crear un data frame con los valores interpolados y de Kalman
base2_interpolada <- data.frame(Date = merged_df$Date, High_interpolado = serie_interpolada, High_kalman = serie_kalman, Original = merged_df$High)

# Base de precios diarios:
ATT_precio_diario <- data.frame(Date = base2_interpolada$Date, Precios_Diarios = base2_interpolada$High_kalman)


# Agregacion ATT --------------------------------------------------------------

### MENSUAL ###

# Genero una base con solo las columnas que me interesan:
nueva_base <- data.frame(base2_interpolada[, c("Date","High_kalman")])

# Convertir la variable Date a formato fecha
nueva_base$Date <- as.Date(nueva_base$Date)

# Crear una variable con el mes y el año
nueva_base$month <- format(nueva_base$Date, "%Y-%m")

# Agrupar por mes y calcular la media
monthly_avg <- aggregate(nueva_base$High_kalman, by = list(month = nueva_base$month), mean)

# Renombrar las columnas
colnames(monthly_avg) <- c("month", "avg_high_kalman")

### Anual ###

# Convertir la variable Date a formato fecha
nueva_base$Date <- as.Date(nueva_base$Date)

# Crear una variable con el año
nueva_base$year <- format(nueva_base$Date, "%Y")

# Agrupar por año y calcular la media
annual_avg <- aggregate(nueva_base$High_kalman, by = list(year = nueva_base$year), mean)

# Renombrar las columnas
colnames(annual_avg) <- c("year", "avg_high_kalman")


# ATT_precio_diario: Base AT&T de precios diarios. - Interpolados con filtro kalman
# 
# monthly_avg: Base PRECIO DE ACCIONES AT&T: mensual, interpolada con filtro kalman
# 
# annual_avg: Base PRECIO DE ACCIONES AT&T: ANUAL, INTERPOLARA CON FILTRO Kalman
# 
# *NOTA: Para almacenarlos sera mejor tener los diarios, anuales y mensuales (precio, acciones y valor) en un mismo data frame



# Numero de acciones AT&T - Diario ----------------------------------------

# AT&T
# 
# Serie con el valor de 1950: 27,585,607.
# 
# En 1950 el promedio de acciones en circulacion era de 27,585,607. (pagina 28)
# 
# ¿Cuantos stock splits hubo? 4
# 
# 1959 - Abril 24 - 3 for 1 common
# 
# 1964 - Mayo 28 - 2 for 1 common
# 
# 1999 - Abril 15 - 3 for 2 common
# 
# 2002 - Noviembre 18 - 1 for 5 reverse
# 
# Para el numero de acciones de ATT obtuve los diarios primero apartir del reporte de 1950 y los splits. Mensualize para obtener mensuales y anuales.


ATT.d <- data.frame(date = ATT_precio_diario$Date)


# Agrego la columna acciones con los cambios de las mismas:
ATT.d <- ATT.d %>%
  mutate(acciones = ifelse(date <= as.Date("1959-04-24"), 27585607,
                           ifelse(date >= as.Date("1959-04-25") & date <= as.Date ("1964-05-27"), 27585607 * 3,
                                  ifelse(date >= as.Date("1964-05-28") & date <= as.Date ("1999-04-15"), (27585607 * 3)* 2,
                                         ifelse(date >= as.Date("1999-04-16") & date <= as.Date ("2002-09-01"), ((27585607 * 3)* 2)*(3/2), ((((27585607 * 3)* 2)*(3/2))*(1/5)))))))



# ATT.d: Base diaria del numero de acciones AT&T


# Numero de acciones AT&T - Mensual-Anual ---------------------------------

# MENSUAL #

# Calcular el promedio mensual utilizando la función aggregate (guardandolo en un nuevo data frame)
ATT.m <- aggregate(acciones ~ format(date, "%Y-%m"), data = ATT.d, mean)

# Corregimos los nombres de columnas:
colnames(ATT.m) <- c("Date", "acciones mensuales")


# ANUAL #

# Calcular el promedio anual utilizando la función aggregate
ATT.a <- aggregate(acciones ~ format(date, "%Y"), data = ATT.d, mean)

# Corregimos los nombres de columnas:
colnames(ATT.a) <- c("Date", "acciones promedio anual")


# ATT.m: Base mensual del numero de acciones AT&T
# 
# ATT.a: Base anual del numero de acciones AT&T



# ¿Como estan ordenados las bases de datos? -------------------------------

#   AT&T:
#   
#   Numero de acciones
# 
# ATT.d: Base diaria del numero de acciones AT&T
# 
# ATT.m: Base mensual del numero de acciones AT&T
# 
# ATT.a: Base anual del numero de acciones AT&T
# 
# Precio de acciones
# 
# ATT_precio_diario: Base AT&T de precios diarios. - Interpolados con filtro kalman
# 
# monthly_avg: Base PRECIO DE ACCIONES AT&T: mensual, interpolada con filtro kalman
# 
# annual_avg: Base PRECIO DE ACCIONES AT&T: ANUAL, INTERPOLARA CON FILTRO Kalman
# 
# AMX:
#   
#   Numero de acciones
# 
# AMX.a.a: Numero de acciones - AMX - Anual
# 
# AMX.a.m.litterman: Numero de acciones - AMX - Mensual
# 
# Precio de acciones
# 
# AMX.d.d : Precio de acciones - AMX -Diario
# 
# AMX.d.m: Precio de acciones - AMX - Mensual
# 
# AMX.d.a: Precio de acciones - AMX - Anual


# Union de dataframes -----------------------------------------------------

# AT&T DIARIOS:
ATT.diario <- data.frame(Date = ATT.d$date, Precio = ATT_precio_diario$Precios_Diarios, Acciones = ATT.d$acciones)

# AT&T Mensual:
ATT.mensual <- data.frame(Date = ATT.m$Date, Precio = monthly_avg$avg_high_kalman, Acciones = ATT.m$acciones)

# AT&T Anual:
ATT.anual <- data.frame(Date = ATT.a$Date, Precio = annual_avg$avg_high_kalman, Acciones = ATT.a$acciones)

# AMX DIARIO:
AMX.diario <- AMX.d.d

# AMX MENSUAL: NOTA: solo obtuve el num de acciones hasta 2021

# Recorto los datos:
# Seleccionar filas despues de 2001-03-01
AMX.a.m.litterman <- subset(AMX.a.m.litterman, Date >= as.Date("2001-03-01"))

# Seleccionar las filas antes de 2021-12-01
AMX.d.m <- subset(AMX.d.m, Date <= as.Date("2021-12-01"))

AMX.mensual <- data.frame(Date = AMX.d.m$Date, Precio = AMX.d.m$AMX.Close, Acciones = AMX.a.m.litterman$Acciones.mensuales)

# AMX Anual:
# Cortar los ultimos dos años de datos
AMX.d.a <- AMX.d.a %>% slice(1:(n() - 2))

AMX.anual <- data.frame(Date = AMX.d.a$`year(Date)`, Precio =AMX.d.a$AMX.Close, Acciones = AMX.a.a$`Promedio ponderado de acciones en circulacion(millones)`)


# Ahora los dataframes completos son:
#   
#   ATT.diario
# 
# ATT.mensual
# 
# ATT.anual
# 
# AMX.diario (incompleto)
# 
# AMX.mensual
# 
# AMX.anual


# Almacenamiento ----------------------------------------------------------
############# Solo sirve para xlsx ya existentes.########

# Define the name and path of the Excel file
filename <- "datos/DATOS.xlsx"

# Create a new workbook and add sheets
wb <- createWorkbook()
addWorksheet(wb, "Daily_AMX")
addWorksheet(wb, "Monthly_AMX")
addWorksheet(wb, "Annual_AMX")
addWorksheet(wb, "Daily_AT&T")
addWorksheet(wb, "Monthly_AT&T")
addWorksheet(wb, "Annual_AT&T")

# Write the data to the sheets

writeData(wb, sheet = "Daily_AMX", x = AMX.diario, startRow = 1, startCol = 1, colNames = TRUE)
writeData(wb, sheet = "Monthly_AMX", x = AMX.mensual, startRow = 1, startCol = 1, colNames = TRUE)
writeData(wb, sheet = "Annual_AMX", x = AMX.anual, startRow = 1, startCol = 1, colNames = TRUE)


writeData(wb, sheet = "Daily_AT&T", x = ATT.diario, startRow = 1, startCol = 1, colNames = TRUE)
writeData(wb, sheet = "Monthly_AT&T", x = ATT.mensual, startRow = 1, startCol = 1, colNames = TRUE)
writeData(wb, sheet = "Annual_AT&T", x = ATT.anual, startRow = 1, startCol = 1, colNames = TRUE)

# Volumen: # lO HAGO MANUAL


# Save the workbook to a new Excel file
saveWorkbook(wb, filename, overwrite = TRUE)





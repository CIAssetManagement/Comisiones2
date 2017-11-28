library(dplyr)
library(tidyr)
library(readxl)
library(xlsx)
library(FundTools)
options(scipen = 999)

###############################################################################################################
#                                           Dias del mes                                                      #
###############################################################################################################

dias <- diff(seq(as.Date("2017-01-01"), Sys.Date()+16, by = "month"))
dias_mes <- as.integer(dias[length(dias)])

###############################################################################################################
#                                               Datos                                                         #
###############################################################################################################

datos <- read.csv("datos2.csv",header = TRUE)
datos <- cbind(datos,Codigo=paste0(datos$emisora," ",datos$serie))
datos$Codigo <- gsub( "\\+", "", datos$Codigo)
datos$TV <- ifelse(datos$emisora == "NAVIGTR",52,substr(as.character(datos$id),start = 1,stop = 2))

comisiones <- read_xlsx("comisiones.xlsx",sheet = "ComisionesFondos")
emisora <- c()
for(i in seq(1,length(comisiones$Fondo),1)){emisora <- c(emisora,strsplit(comisiones$Fondo[i]," ")[[1]][1])}
comisiones$Emisora <- emisora
tv52 <- c("AXESEDM","NAVIGTR","CIBOLS","CIEQUS")
comisiones$TV <- ifelse(comisiones$Emisora %in% tv52,52,51)
comisiones$id <- comisiones$TV
comisiones$id <- ifelse(substr(gsub(" ","-",comisiones$Fondo),1,1) == "C",
                        paste0(comisiones$id,"-","+",gsub(" ","-",comisiones$Fondo)),
                        paste0(comisiones$id,"-",gsub(" ","-",comisiones$Fondo)))
comisiones$Emisora <- NULL
comisiones$TV <- NULL

###############################################################################################################
#                                               Títulos                                                       #
###############################################################################################################

titulos <- datos %>% group_by(Codigo,fecha,TV) %>% summarise(Titulos=sum(tit))
titulos <- titulos %>% filter(Codigo %in% comisiones$Fondo)
titulos$Codigo <- ifelse(substr(titulos$Codigo,start = 1,stop = 1)=="C",paste0("+",titulos$Codigo),titulos$Codigo)
titulos$id <- paste0(titulos$TV,"-",gsub(" ","-",titulos$Codigo))
titulos$TV <- NULL
titulos$fecha <- as.Date(titulos$fecha,format='%Y-%m-%d')

###############################################################################################################
#                                               Precios                                                       #
###############################################################################################################

precios <- get_prices(titulos$fecha,titulos$id)
precios[is.na(precios)] <- 0
precios$fecha <- as.Date(precios$fecha,format='%Y-%m-%d')

###############################################################################################################
#                                     Completando los días perdidos                                           #
###############################################################################################################

#Habiles para completar títulos de forma adecuada.
habiles <- unique(precios$fecha)

#Fechas para hacer el complete de precios
fechas <- seq.Date(titulos$fecha[1],titulos$fecha[length(titulos$fecha)],1)
precios <- precios %>% complete(fecha = fechas)

#Poniendo precios para todas las fechas del mes
for(i in seq(1,length(precios$fecha),1)){
  for(j in seq(2,length(colnames(precios)),1)){
    if(is.na(precios[i,j]) == TRUE)
      precios[i,j] <- precios [i-1,j]
  }
}

#Completando títulos y poniendo los títulos para todas las fechas del mes
titulos <- titulos %>% complete(nesting(Codigo,id),fecha = fechas)
titulos <- titulos[!duplicated(titulos), ]
inhabiles <- seq(Sys.Date(),length = 31,by=-1)
inhabiles <- inhabiles[!(inhabiles %in% habiles)]
for(i in seq(1,length(titulos$fecha),1)){
  if(is.na(titulos$Titulos[i]) == TRUE){
    if(titulos$fecha[i] %in% inhabiles)
      titulos$Titulos[i] <- titulos$Titulos[i-1]
    else
      titulos$Titulos[i] <- 0
  }
}

###############################################################################################################
#                                       Calculando los Montos del mes                                         #
###############################################################################################################

montos <- precios
for(i in seq(2,length(montos$fecha)-1,1)){
  for(j in seq(2,length(colnames(montos)),1)){
    primerindice <- which(titulos$id == colnames(montos)[j])
    segundoindice <- which(titulos$fecha == (montos$fecha[i])-1)
    indice <- intersect(primerindice,segundoindice)
    cindice <- which(comisiones$id == colnames(montos)[j])
    if(length(indice) == 0)
      montos[i,j] <- 0
    else
      montos[i,j] <- precios[i+1,j]*titulos$Titulos[indice]*comisiones$Comision[cindice]/365
  }
}

for(i in seq(1,length(inhabiles),1)){
  if(montos$fecha[i] %in% inhabiles)
    montos[i,-1] <- montos[i-1,-1]
}

diahoy <- as.numeric(format(as.Date(Sys.Date()-2,format="%Y-%m-%d"), "%d"))
fechasmes <- seq(Sys.Date()-2,length = diahoy,by=-1)
montos <- montos[montos$fecha %in% fechasmes,]
total <- data.frame(rowSums(montos[-1]))
colnames(total) <- c("Monto Total del Día")
rownames(total) <- montos$fecha
total$`Monto Total del Día` <- format(round(total$`Monto Total del Día`,digits=2), big.mark=",", scientific=FALSE)
total$`Monto Total del Día` <- paste0("$",total$`Monto Total del Día`)

###############################################################################################################
#                                            Escribiendo el Excel                                             #
###############################################################################################################
comisiones <- data.frame(comisiones[,c(1,2)])

titulos1 <- titulos
titulos <- precios
for(i in seq(2,length(precios$fecha),1)){
  for(j in seq(2,length(colnames(precios)),1)){
    primerindice <- which(titulos1$id == colnames(precios)[j])
    segundoindice <- which(titulos1$fecha == (precios$fecha[i])-1)
    indice <- intersect(primerindice,segundoindice)
    titulos[i,j] <- titulos1$Titulos[indice]
  }
}
titulos <- data.frame(titulos)
titulos <- titulos[titulos$fecha %in% fechasmes,]

precios <- data.frame(precios)
precios <- precios[precios$fecha %in% fechasmes,]

dia <- Sys.Date()-1
archivom <- paste0('C:/Users/MATREJO/Documents/Comisiones/fondosCISM_',dia,'.xlsx')
write.xlsx2(x = total,file = archivom,sheetName = "MontosDiarios",col.names = TRUE,row.names = TRUE,append = FALSE)
write.xlsx2(x = comisiones,file = archivom,sheetName = "Comisiones",col.names = TRUE,row.names = FALSE,append = TRUE)
write.xlsx2(x = precios,file = archivom,sheetName = "Precios",col.names = TRUE,row.names = FALSE,append = TRUE)
write.xlsx2(x = titulos,file = archivom,sheetName = "Titulos",col.names = TRUE,row.names = FALSE,append = TRUE)

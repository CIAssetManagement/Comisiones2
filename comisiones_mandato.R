library(dplyr)
library(tidyr)
library(readxl)
library(xlsx)

###############################################################################################################
#                                           Dias del mes                                                      #
###############################################################################################################

dias <- diff(seq(as.Date("2017-01-01"), Sys.Date()+16, by = "month"))
dias_mes <- as.integer(dias[length(dias)])
#Cambiar el -1 por el último dia del que se tengan datos, es decir, si son de antier va un -2.
diahoy <- Sys.Date() - 3

###############################################################################################################
#                                               Datos                                                         #
###############################################################################################################

hoy <- Sys.Date()
mes <- substr(hoy,6,7)
ao <- substr(hoy,1,4)
datos <- get_position(seq_Date(paste0(ao,mes,"01/")))
datos <- cbind(datos,Codigo=paste0(datos$emisora," ",datos$serie))
datos$Codigo <- gsub( "\\+", "", datos$Codigo)

###############################################################################################################
#                                         Cajón en el mandato                                                 #
###############################################################################################################

cajon <- function(monto){
  montos <- c(999999,5000000,7500000,10000000,30000000,100000000)
  if(monto < montos[1]){renglon <- 1}
  if(between(monto,montos[1],montos[2])){renglon <- 2}
  if(between(monto,montos[2],montos[3])){renglon <- 3}
  if(between(monto,montos[3],montos[4])){renglon <- 4}
  if(between(monto,montos[4],montos[5])){renglon <- 5}
  if(between(monto,montos[5],montos[6])){renglon <- 6}
  if(monto > montos[6]){renglon <- 7}
  return(renglon)
}

###############################################################################################################
#                                        Carteras Testigo                                                     #
###############################################################################################################

lista <- data.frame(read_excel("comisiones.xlsx",sheet = 'Testigo'))
testigo <- ifelse(datos$contrato %in% lista$contrato, "Sacar", "No Sacar")
datos <- cbind(datos,testigo)
datos <- datos %>% filter(testigo == "No Sacar")
datos$testigo <-  NULL

###############################################################################################################
#                                        Posiciones sin fondos                                                #
###############################################################################################################

lista <- c("AXESEDM","NAVIGTR","+CIGUB","+CIGUMP","+CIGULP","+CIPLUS","+CIUSD","+CIEQUS","+CIBOLS")
funds <- ifelse(datos$emisora %in% lista, "Fondo", "No Fondo")
datos <- cbind(datos,funds)

sinfondos <- datos %>% filter(funds == "No Fondo")
#La siguiente linea se usa en caso de cobrar con mandato lo que pueda estar en fondos.
#sinfondos <- datos
sinfondos <- data.frame(fecha =sinfondos$fecha,contrato =sinfondos$contrato,cartera = sinfondos$carteramodelo ,
                        mon = sinfondos$mon,Codigo =sinfondos$Codigo,Emisora = sinfondos$emisora)

sinfondos$mon <- ifelse(sinfondos$mon > 0 , sinfondos$mon, 0)

###############################################################################################################
#                                                 Grupos                                                      #
###############################################################################################################

grupos <- data.frame(read_excel("comisiones.xlsx",sheet = 'Grupos'))
indicesg <- match(sinfondos$contrato,grupos$Titular)
grupo <- grupos$Grupo[indicesg]
sinfondos <- cbind(sinfondos,Grupo = grupo)
grupo1 <- sinfondos %>% group_by(Grupo,fecha) %>% summarise(Monto_Grupal = sum(mon))
sinfondos <- data.frame(merge(sinfondos, grupo1, by.x=c("fecha", "Grupo"), by.y=c("fecha", "Grupo")))
indiceserror <- which(is.na(sinfondos$Grupo) == TRUE)
sinfondos$Monto_Grupal[indiceserror] <- sinfondos$mon[indiceserror]

###############################################################################################################
#                                             Clientes especiales                                             #
###############################################################################################################

especiales <- data.frame(read_excel("comisiones.xlsx",sheet = 'Especiales'))
indicese <- match(sinfondos$contrato,especiales$Titular)
descuento <- 1 - especiales$Descuento[indicese]
descuento[is.na(descuento)] <- 1
sinfondos <- cbind(sinfondos,Descuento=descuento)

###############################################################################################################
#                              Comisiones por instrumentos fuera de fondos                                    #
###############################################################################################################

comision_sinfondo <- data.frame(read_excel("comisiones.xlsx",sheet = 'ComisionesSinFondos'))
colnames(comision_sinfondo) <- gsub("\\.","-",colnames(comision_sinfondo))
colnames(comision_sinfondo) <- gsub("\\_"," ",colnames(comision_sinfondo))

###############################################################################################################
#                                           Saldo Diario                                                      #
###############################################################################################################

saldo_diario <- sinfondos %>% group_by(contrato,fecha) %>% summarise(Monto = sum(mon))
fechas <- unique(saldo_diario$fecha)
saldo_diario <- saldo_diario %>% complete(nesting(contrato),fecha = fechas,fill = list(Monto = 0))
diario <- data.frame(matrix(saldo_diario$Monto,nrow = length(fechas)),row.names = fechas)
colnames(diario) <- unique(saldo_diario$contrato)

###############################################################################################################
#                                           Cobros Especiales                                                 #
###############################################################################################################

cespeciales <- data.frame(read_excel("comisiones.xlsx",sheet = 'CobrosEspeciales'))
cespeciales$Contrato <- as.integer(cespeciales$Contrato)
bonificacion <- rep(0,length(sinfondos$contrato))
for (i in seq(1,length(cespeciales$Contrato), 1)){
  indicesb <- match(cespeciales$Contrato[i],sinfondos$contrato)
  bonificacion[indicesb] <- cespeciales$Bonificacion[i]
}
sinfondos <- cbind(sinfondos,Bonificacion = bonificacion)

###############################################################################################################
#                                               Cobro                                                         #
###############################################################################################################

saldopromedio <- colMeans(diario)
contratos <- colnames(diario)

cartera <- as.character(sinfondos$cartera[match(contratos,sinfondos$contrato)])
cobro <- data.frame(cbind(Contrato = contratos,SaldoProm = saldopromedio,Cartera = cartera))
cobro$SaldoProm <- as.numeric(as.character(cobro$SaldoProm))
Datoscobro <- sinfondos %>% group_by(contrato) %>% summarise(Promediogrupal = sum(Monto_Grupal)/as.numeric(length(unique(sinfondos$fecha))), 
                                                             Descuento = mean(1-Descuento),
                                                             Bonificacion = sum(Bonificacion))

#Comision anual
indicescolumna <- match(cartera,colnames(comision_sinfondo))
indicesrenglon <- as.integer(sapply(Datoscobro$Promediogrupal,cajon))
comision_anuals <- c()
for (i in seq(1,length(indicescolumna),1)){
  comision_anuals <- c(comision_anuals,comision_sinfondo[indicesrenglon[i],indicescolumna[i]])
}

Datoscobro <- cbind(Datoscobro,Anual = comision_anuals)

comision <- dias_mes*Datoscobro$Anual/360
comisiondesc <- comision * (1-Datoscobro$Descuento)
#############
monto <- (cobro$SaldoProm * dias_mes*Datoscobro$Anual/360)
monto <- ifelse(Datoscobro$Bonificacion <= 0, monto,Datoscobro$Bonificacion/1.16)
#############
montodesc <- (saldopromedio * comisiondesc)
montodesc <- ifelse(montodesc < 0, 0, montodesc)
montodesc <- ifelse(Datoscobro$Bonificacion <= 0, montodesc,Datoscobro$Bonificacion/1.16)
montoiva <- montodesc * 1.16
montoneto <- ifelse(Datoscobro$Bonificacion <= 0, montodesc*1.16,Datoscobro$Bonificacion)
cobro <- cbind(cobro,Anual = Datoscobro$Anual,Comision = comision, Descuento = Datoscobro$Descuento,
               ComisionConDescuento = comisiondesc, Monto = monto, MontoconDescuento = montodesc, 
               MontoIVA = montoiva, Bonificacion = Datoscobro$Bonificacion,MontoNetoIVA = montoneto, 
               CobroFinal = montoneto)

cobro[-c(1,3)] <- cobro[-c(1,3)] %>% lapply(as.character) %>% lapply(as.numeric)

###############################################################################################################
#                                             Cargos                                                          #
###############################################################################################################

cargos  <- data.frame(cbind(Contrato = as.character(cobro$Contrato), CobroFinal = cobro$CobroFinal))
cargos$CobroFinal <- as.numeric(as.character(cobro$CobroFinal)) %>% round(4)

###############################################################################################################
#                                   Efectivo que no está en corto                                             #
###############################################################################################################

sinfondos$Codigo <- as.character(sinfondos$Codigo)
sinfondos$fecha <- as.character(sinfondos$fecha)

suficientes <- sinfondos %>% filter(fecha == as.character(diahoy)) %>% filter(Codigo == "EFECTIVO ")
suficientes <- data.frame(cbind(Contrato = suficientes$contrato,Efectivo =suficientes$mon))
efectivo <- data.frame(merge(cargos,suficientes,by.x=c('Contrato'),by.y=c('Contrato'),all.x = TRUE))
efectivo$Efectivo <- ifelse(is.na(efectivo$Efectivo)==TRUE,0,efectivo$Efectivo)
efectivo$Contrato <- as.character(efectivo$Contrato)
efectivo$Efectivo <- ifelse(efectivo$Efectivo>0,efectivo$Efectivo,0)
efectivosuficiente <- sum(ifelse(efectivo$Efectivo > efectivo$Cobro,efectivo$Cobro,efectivo$Efectivo))

###############################################################################################################
#                                              Resumen                                                        #
###############################################################################################################

nomresumen <- c('Dias en el mes','Suma sin Descuentos','Suma neta sin IVA','Suma neta con IVA',
                'Suma neta con IVA quitando fondos insuficientes')
resumen <- data.frame(matrix(c(dias_mes,sum(cobro$Monto), sum(cobro$MontoconDescuento),sum(cobro$MontoNetoIVA), 
                               efectivosuficiente), nrow = 5),row.names = nomresumen)

###############################################################################################################
#                                                Excel                                                        #
###############################################################################################################

#Comisiones por mandato
dia <- Sys.Date()
archivom <- paste0('C:/Users/MATREJO/Documents/Comisiones/cargosCISM_',dia,'.xlsx')
write.xlsx2(cargos,archivom, sheetName = 'cargosCISM', col.names=TRUE, row.names = FALSE, append = FALSE)
write.xlsx2(resumen,archivom, sheet = 'ResumenCobro', col.names = FALSE, row.names = TRUE, append = TRUE)
write.xlsx2(cobro,archivom, sheet = 'completocobro', col.names = TRUE, row.names = FALSE, append = TRUE)
write.xlsx2(diario,archivom, sheet = 'saldoDiarioContratosMes', col.names = TRUE, row.names = TRUE,
            append = TRUE)
write.xlsx2(grupos,archivom, sheet = 'Grupos', col.names = TRUE, row.names = FALSE, append = TRUE)
write.xlsx2(especiales,archivom, sheet = 'DescuentosEmpleados', col.names = TRUE, row.names = FALSE, 
            append = TRUE)
write.xlsx2(cespeciales,archivom, sheet = 'BonificacionesCargos', col.names = TRUE, row.names = FALSE, 
            append = TRUE)
#Clientes en corto
archivon <- 'C:/Users/MATREJO/Documents/Comisiones/ClientesenCorto.xlsx'
write.xlsx2(efectivo,archivon, sheetName = 'Clientes', col.names=TRUE, row.names = FALSE, append = FALSE)

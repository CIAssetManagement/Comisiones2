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
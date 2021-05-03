library(openxlsx)

#Preparacion archivo para modelo linear mixto

nombre_archivos=list.files(path="D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel")

nombre_archivos=nombre_archivos[!nombre_archivos %in% "Calculado_Incremento_BAI"]



for(i in c(1:length(nombre_archivos))){
  nombre_archivos[i]=substr(nombre_archivos[i], 1, 7)
}


contador_filas=1

for(n in nombre_archivos){

  x=read.xlsx(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI/",n,".xlsx",sep=""), colNames=TRUE,rowNames=TRUE)
  nombre=colnames(x)
  
  dimension=dim(x)
  fil=dimension[1]
  col=dimension[2]
  
  matriz=matrix(nrow = fil*col, ncol = 11)
  
  anio=rep(rownames(x),times=col)
  
  matriz[,2]=anio
  
  localizaciones=read.xlsx("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/site_details.xlsx")
  localizaciones=localizaciones[contador_filas,c(3:6)]
  
  matriz[,3]=rep(localizaciones[[1]])
  matriz[,4]=rep(localizaciones[[2]])
  matriz[,5]=rep(localizaciones[[3]])
  matriz[,6]=rep(localizaciones[[4]])
  
  contador=0
  
  matriz<- as.data.frame(matriz)
  
  for (i in c(1:col)){
    edad=length(which(!is.na(x[[i]])))
    for (j in c(1:fil)){
      contador=contador+1
      matriz[[contador,1]]=contador
      matriz[[contador,7]]=nombre[[i]]
      matriz[[contador,8]]=edad-(fil-j)
      matriz[[contador,11]]=x[[j,i]]
    }
  }
  
  for(o in c(4:6,8:11)){
    matriz[,o]=as.double(matriz[,o])
  }

  
  for(i in c(2:(col*fil))){
    if(i!=(col*fil)){
      matriz[i,9]=((pi*(matriz[[i,11]]^2))-(pi*(matriz[[i-1,11]])^2))/100
      matriz[i+1,10]=matriz[[i,9]]
    }else {
      matriz[i,9]=((pi*(matriz[[i,11]]^2))-(pi*(matriz[[i-1,11]])^2))/100
    }
  }
  
  colnames(matriz)=c("n","year","sp","el","lat","lon","tree","age","bai","baip","dbh")
  
  write.xlsx(matriz,paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI/Archivos_MLM_JuanCarlos/",n,".xlsx",sep=""),row.names = TRUE,col.names = TRUE,showNA = FALSE)

  contador_filas=contador_filas+1
}


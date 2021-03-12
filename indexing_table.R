#####################################################
##### Indexing file - prepare table from matrix #####
#####################################################

# install needed packages
install.packages(c("XLConnect","stringr"))

library(XLConnect)
library(stringr)


##### TEST ######

setwd("C:/Users/Mateusz/Uniwersytet Jagielloñski/Methods - Dokumenty/Sequencing/Scripts")
getwd() # check if path is ok

MiSeq_sheet <- function(file){
  
  indexing <- loadWorkbook(file)
  sheetNames <- getSheets(indexing)
  
  index_table <- data.frame(matrix(ncol = 4, nrow = 0))
  
  for (i in sheetNames) {
    first <- readWorksheet(indexing, sheet=i)
    first[,"last_column"] <- NA
    k <- which(first == "plate", arr.ind = TRUE)
    
    if (length(k[,1]) == 0){
      
      next()
      
    } else {
      
      for (j in 1:length(k[,1])) {
        
        row <- k[j,1]
        col <- k[j,2]
        col_1 <- col
        
        while (!is.na(first[row,col])) {
          
          col = col+1
        }
        
        plate <- first[row:(row+9),col_1:(col-1)]
        
        i5 <- as.vector(as.matrix(plate[3:10,1]))
        i5 <- rep(i5, length(plate[1,])-2)
        
        i7 <- as.vector(as.matrix(plate[1,3:length(plate[1,])]))
        i7 <- rep(i7, each=8)
        
        samples <- as.vector(as.matrix(plate[-c(1,2),-c(1,2)]))
        
        plate <- plate[2,2]#i
        
        table_fragm <- cbind(i5, i7, samples)
        table_fragm <- cbind(table_fragm, plate)
        index_table <- rbind(index_table, table_fragm)
        
      }
      
    }
    
  }
  
  exc_new <- loadWorkbook(paste("new",file),create= TRUE) # creates new xls under old name + "new"
  createSheet(exc_new, name= substr(i, 1,8))
  writeWorksheet(exc_new, index_table, sheet=1) 
  saveWorkbook(exc_new,paste("./",paste0("MiSeq_",file)))
}


MiSeq_sheet("Greeenland_example.xlsx") # in the brackets write name of your file

rm(list=ls())



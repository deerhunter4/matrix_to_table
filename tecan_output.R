
##################################################
##### Tecan output files transformation code #####
##################################################

# first set your working directory (wd), 
# so give path to the folder where you store file that you want to transform
# remember to analayze files from each week separately

setwd("/Users/User/Dropbox/R_scripts/sunrise/sunrise_20171108")
setwd("/Users/Mateusz/Dropbox/R_scripts/sunrise/sunrise_20171108")
getwd() # check if path is ok

# create folders in which intermediate and final files will we sored
dir.create("./new/plates/strain", recursive=T)

# install needed packages
install.packages(c("XLConnect","stringr"))

library(XLConnect)
library(stringr)


##################
##### Step I #####


fileNames <- Sys.glob("*.xls") # it puts names of all excel files from working directory to "fileNames"
fileNames

## run all lines belowe, from "Tecan_1" to last "}" to load function "Tecan_1"

Tecan_1 <- function(fileNames){
  
  for (i in fileNames) {
    file_1 <- loadWorkbook(i) # loading first xls file into "file_1"
    sheetNames <- getSheets(file_1) # gets sheets names from file_1
    sheetNames<- sheetNames[1:length(sheetNames)-1] # delete last sheet, which is empty
    day_1 <- data.frame(matrix(0, ncol = 1+length(sheetNames), nrow = 96))
    day_1 <- setNames(day_1, c("ID",sheetNames)) # creates dataframe with columns names corresponding with sheets names
    ID<- expand.grid(LETTERS[1:8],1:12) #
    ID<- paste(ID[,1],ID[,2],sep="")    #
    day_1[,"ID"] <- ID                  # puts Tecan plate cell numbers A1-H12 to the first column
    for (j in sheetNames) {
      first <- readWorksheet(file_1, sheet=j) # reads first sheet from file_1 as data frame
      first<- first[1:8,2:13] # extract part of data frame containing Tecan measures
      first<- as.matrix(first)
      first<- as.vector(first) # create a vector from matrix
      first<- as.numeric(gsub(",", ".", first)) # replace comma with dot
      first<- round(first,4)
      day_1[,j] <- first
    }
    exc_new <- loadWorkbook(paste("new",i),create= TRUE) # creates new xls under old name + "new"
    createSheet(exc_new, name= substr(i, 1,8))
    writeWorksheet(exc_new, day_1, sheet=1) # writes to the created sheet data frame "day_1" that contain all Tecan measures for all C. elegans mutants (mutant names as column names)
    saveWorkbook(exc_new,paste("./new/",i))
  }
}


ls() # you can check what elements you have in your working environment


Tecan_1(fileNames) # running function Tecan_1 on files from your working directory

# check files in "new" folder


#rm(list=ls())


###################
##### Step II #####

#change wd to "new" folder, to use newly created files

setwd("/Users/Mateusz/Dropbox/R_scripts/sunrise/sunrise_20171108/new")


fileNames <- Sys.glob("*.xls") # it puts names of all excel files from working directory to "fileNames"
fileNames

# load function "Tecan_2"

Tecan_2 <- function(fileNames){
  plates_names <- vector(mode="numeric", length=0)
  for (i in fileNames) {
    file_1 <- loadWorkbook(i) # loading first xls file into "file_1"
    first <- readWorksheet(file_1, sheet=1) # reads first sheet from file_1 as data frame
    plates_1 <- colnames(first)
    plates_names <- c(plates_names,plates_1)
  }
  
  plates_names <- unique(plates_names) 
  plates_names <- plates_names[-1] # I do not know why, but R read ";" as dot in column names
  
  for (j in plates_names){
    day_1 <- data.frame(matrix(0, ncol = 2, nrow = 96))
    day_1 <- setNames(day_1, c("ID",j)) # creates dataframe with columns names corresponding with sheets names
    ID<- expand.grid(LETTERS[1:8],1:12) #
    ID<- paste(ID[,1],ID[,2],sep="")    #
    day_1[,"ID"] <- ID    
    day_1[,j]<- rep(j,96)
    trial<- as.data.frame(do.call(rbind, str_split(day_1[,j],"\\."))) # "\\." is needed to divide string in dots
    colnames(trial) <- as.character(unlist(trial[1,]))
    day_1<- cbind(day_1,trial)
    day_1<- day_1[,-2]
    colnames(day_1)<- c("ID","Plate","treatment","Strain","number") # remember to give names of columns in proper order
    
    for(i in fileNames){
      file_1 <- loadWorkbook(i) 
      first <- readWorksheet(file_1, sheet=1)
      
      if (is.element(j, colnames(first))) {
        first<- first[,j]
        day_1<- cbind(day_1,first)
        colnames(day_1)[length(day_1[1,])]<- substr(i, 1,9)
        day_1
      }
      
      else {
        next()
      }
    }
    exc_new <- loadWorkbook(paste("new",j, ".xls"),create= TRUE) 
    createSheet(exc_new, name= j)
    writeWorksheet(exc_new, day_1, sheet=1) # writes to the created sheet data frame "day_1" that contain all Tecan measures for all C. elegans mutants (mutant names as column names)
    saveWorkbook(exc_new,paste("./plates/",j, ".xls"))
  }
  
  
}


Tecan_2(fileNames) 


####################
##### Step III #####

setwd("/Users/Mateusz/Dropbox/R_scripts/sunrise/sunrise_20171108/new/plates")


fileNames <- Sys.glob("*.xls") # it puts names of all excel files from working directory to "fileNames"
fileNames


Tecan_3 <- function(fileNames){
  
  IDs<- c("A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "A2", "B2", 
          "C2", "D2", "E2", "F2", "G2", "H2", "A3", "B3", "C3", "D3", "E3", 
          "F3", "G3", "H3", "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", 
          "A5", "B5", "C5", "D5", "E5", "F5", "G5", "H5", "A6", "B6", "C6", 
          "D6", "E6", "F6", "G6", "H6", "A7", "B7", "C7", "D7", "E7", "F7", 
          "G7", "H7", "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "A9", 
          "B9", "C9", "D9", "E9", "F9", "G9", "H9", "A10", "B10", "C10", 
          "D10", "E10", "F10", "G10", "H10", "A11", "B11", "C11", "D11", 
          "E11", "F11", "G11", "H11", "A12", "B12", "C12", "D12", "E12", 
          "F12", "G12", "H12")
  
  bac<-c("pzf-1", "egl-27", "csb-1", "spas-1", "msh-2", "brc-1", "trpa-1", 
         "umps-1", "slx-1", "mev-1", "cep-1", "hsp-3", "cyp-14A5", "T7", 
         "egl-8", "polk-1", "xpf-1", "klp-15", "shc-1", "him-5", "old-1", 
         "ucp-4", "him-8", "egl-8", "cyp-14A5", "umps-1", "hsr-9", "ercc-1", 
         "daf-16", "brd-1", "sgo-1", "cep-1", "cpna-2", "xpg-1", "ife-2", 
         "xpf-1", "him-4", "ire-1", "brc-1", "fcd-2", "pms-2", "brd-1", 
         "polh-1", "rde-2", "mus-81", "mek-1", "natc-1", "nsy-1", "daf-16", 
         "ife-2", "mev-1", "csb-1", "him-8", "fbf-2", "exo-1", "him-4", 
         "rde-2", "hsp-4", "xpa-1", "cpna-2", "polk-1", "egl-27", "fan-1", 
         "shc-1", "age-1", "fan-1", "fbf-2", "pzf-1", "spas-1", "slx-1", 
         "mus-81", "natc-1", "klp-15", "him-5", "age-1", "hsr-9", "trpa-1", 
         "fcd-2", "ucp-4", "ire-1", "polh-1", "msh-2", "abl-1", "hsp-4", 
         "nsy-1", "mek-1", "pms-2", "T7", "abl-1", "xpg-1", "ercc-1", 
         "old-1", "xpa-1", "exo-1", "hsp-3", "sgo-1")
  
  bac_plates<- data.frame(IDs, bac) # you get combination of well ID (from plate A & B) and bacteria strain
  
  szczepy <- vector(mode="numeric", length=0)
  
  for (i in fileNames) {
    file_1 <- loadWorkbook(i) # loading first xls file into "file_1"
    first <- readWorksheet(file_1, sheet=1) # reads first sheet from file_1 as data frame
    szczep <- first[1,"Strain"]
    szczepy <- c(szczepy, szczep)
  }
  
  szczepy <- unique(szczepy) # all unique strains names of C.elegans from all xls files
  
  for (j in szczepy){
    
    day_1 <- data.frame(matrix(0, ncol = 10, nrow=0)) # if you have measures from more than 5 days, then change "10" by adding additional number of days
    
    for(i in fileNames){
      file_1 <- loadWorkbook(i) 
      first <- readWorksheet(file_1, sheet=1)
      
      if (is.element(j, first[1,"Strain"])) {
        
        if(length(first[1,])==10) { # if you have measures from more than 5 days, then change "10" by adding additional number of days
          day_1<- rbind(day_1,first)
          day_1
        }
        else {
          cat('WARNING!!!\n')
          cat("The number of columns in ", i, " is wrong")
          cat('\n')
        }
      }

      else {
        next()
      }

    }
    #day_1$ID<- with(day_1, paste(ID, Plate, sep="_"))
    day_1<- merge(day_1,bac_plates, by.x="ID", by.y="IDs")
    #day_1<- day_1[,-1]
    day_1<- day_1[,c(11,1:10)] # if you have measures from more than 5 days, then change "10" and "11" by adding additional number of days
    exc_new <- loadWorkbook(paste("new",j, ".xls"),create= TRUE) 
    createSheet(exc_new, name= j)
    writeWorksheet(exc_new, day_1, sheet=1) # writes to the created sheet data frame "day_1" that contain all Tecan measures for all C. elegans mutants (mutant names as column names)
    saveWorkbook(exc_new,paste("./strain/",j, ".xls"))
  }
}


ls()

Tecan_3(fileNames)


rm(list=ls())


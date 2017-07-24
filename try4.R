library(rJava)
library(xlsxjars)
library(xlsx)
getwd()
setwd("C:/Users/Documents/R")
d <- read.xlsx("Port.xlsx",sheetIndex = 3)
d <- d[order(d[,"VESSEL"]),]
dat <- d[!is.na(d$PORT), ]
for(a in 1:17)
{dat[,a] <- as.character(dat[,a])
dat[,a][is.na(dat[,a])] <- "Unknown" }

##Other duplicates

library(dplyr)
dati <- filter(dat,STATUS == "VESSEL AT OUTER ANCHORAGE" | STATUS == "VESSEL AT KUTUBDIA" | STATUS == "VESSEL AT ANCHOARGE" | STATUS=="VESSEL NOT ENTERING" | STATUS==" ")
da<-dati[!duplicated(dati[,c(2,3,6)]),]
da <- da[order(da[,"VESSEL"]),]
##MAIN duplicates
library(plyr)
r <- ddply(dat,.variables = c("PORT","VESSEL","IMP.EXP","QTY..Mts.","RECEIVERS"),
           function(t){if(t$IMP.EXP[1]=="IMP" | t$IMP.EXP[1]=="IMPORT"){
                   t$STATUS<-factor(x = t$STATUS,levels =c("EXPECTED","ANCHORAGE","BERTH","SAILED"),ordered = T)
                   return(t[which.max(as.integer(t$STATUS)),])
           }else{
                   t$STATUS<-factor(x = t$STATUS,levels =c("BERTH","EXPECTED","ANCHORAGE","SAILED"),ordered = T)
                   return(t[which.max(as.integer(t$STATUS)),])}
                   r<- t[which.max(as.integer(t$STATUS)),]
           }
)
##ORDERING VESSEL
r <- r[order(r[,"VESSEL"]),]
##BINDING ROWS
cleaned <- rbind(r,da)
cleaned <- cleaned[order(cleaned[,"VESSEL"]),]

##SAVING DATA TO EXCEL
library(xlsx)
write.xlsx(cleaned,"cleaned.xlsx")




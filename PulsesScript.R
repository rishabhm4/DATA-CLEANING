library(rJava)
library(xlsxjars)
library(xlsx)
getwd()
setwd("C:\\Users\\Level A\\Documents\\Port Lineup")
d <- read.xlsx("Port.xlsx",sheetIndex = 1)

##For reports having NA value in IMP.EXP


d <- d[order(d[,"VESSEL"]),]
dat <- d[!is.na(d$PORT), ]
for(a in 1:17)
{dat[,a] <- as.character(dat[,a])
dat[,a][is.na(dat[,a])] <- "Unknown" }

##Other duplicates

library(dplyr)
dati <- filter(dat,STATUS == "VESSEL AT OUTER ANCHORAGE" | STATUS == "VESSEL AT KUTUBDIA" | STATUS == "VESSEL AT ANCHOARGE" | STATUS=="VESSEL NOT ENTERING" | STATUS==" ")


da<-dati[!duplicated(dati[,c(2,3,9)]),]
da <- da[order(da[,"VESSEL"]),]
##MAIN duplicates
library(plyr)
r <- ddply(dat,.variables = c("PORT","VESSEL","QTY","RECEIVER"),
           function(t){
             t$STATUS<-factor(x = t$STATUS,levels =c("EXPECTED","ANCHORAGE","BERTH","SAGAR","SANDHEADS","DIAMOND","SAILED"),ordered = T)
             return(t[which.max(as.integer(t$STATUS)),])
           
             r<- t[which.max(as.integer(t$STATUS)),]
           }
)
##ORDERING VESSEL
r <- r[order(r[,"VESSEL"]),]
##BINDING ROWS
cleaned <- rbind(r,da)
cleaned <- cleaned[order(cleaned[,"VESSEL"]),]
de <- cleaned[colSums(!is.na(cleaned)) > 0]

##SAVING DATA TO EXCEL

library(xlsx)
write.xlsx(de,"pulses report.xlsx",sheetName = "19-july-2016",append = TRUE)


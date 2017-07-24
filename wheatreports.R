##WHEAT REPORT
library(rJava)
library(xlsxjars)
library(xlsx)
getwd()
setwd("C:\\Users\\Level A\\Documents\\Port Lineup")
d <- read.xlsx("WHEAT POSITION DATED 17.07.17.xls",sheetIndex = 1)

##For reports having NA value in IMP.EXP
library(dplyr)
d <- filter(d,IMP.EXP == "IMP" | IMP.EXP == "IMPORT" | IMP.EXP == "EXP")

d <- d[order(d[,"VESSEL.NAME"]),]
dat <- d[!is.na(d$PORT), ]
for(a in 1:17)
{dat[,a] <- as.character(dat[,a])
dat[,a][is.na(dat[,a])] <- "Unknown" }

##Other duplicates

library(dplyr)
dati <- filter(dat,STATUS == "VESSEL AT OUTER ANCHORAGE" | STATUS == "VESSEL AT KUTUBDIA" | STATUS == "VESSEL AT ANCHOARGE" | STATUS=="VESSEL NOT ENTERING" | STATUS==" ")

##FOR SUGAR REPORTS --2=STATUS, 3=VESSEL, 6=QTY..Mts.
##FOR  WHEAT REPORTS --2=STATUS, 4=VESSEL, 8=QTY..Mts.
da<-dati[!duplicated(dati[,c(2,4,8)]),]
da <- da[order(da[,"VESSEL.NAME"]),]
library(plyr)
r <- ddply(dat,.variables = c("PORT","VESSEL.NAME","IMP.EXP","QTY.MTS","RECEIVER"),
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
r <- r[order(r[,"VESSEL.NAME"]),]
##BINDING ROWS
cleaned <- rbind(r,da)
cleaned <- cleaned[order(cleaned[,"VESSEL.NAME"]),]
de <- cleaned[colSums(!is.na(cleaned)) > 0]
##SAVING DATA TO EXCEL

library(xlsx)
write.xlsx(de,"wheat report.xlsx",sheetName = "17-july-2016",append = TRUE)


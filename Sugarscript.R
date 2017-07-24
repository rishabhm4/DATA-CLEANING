
library(rJava)
library(xlsxjars)
library(xlsx)
getwd()
setwd("C:\\Users\\Dinu Level A\\Desktop\\Documents\\Port Lineup")
d <- read.xlsx("DAILY SUGAR VESSEL POSITION -14.07.16.xls",sheetIndex = 1)

##For reports having NA value in IMP.EXP
library(dplyr)
d <- filter(d,IMP.EXP == "IMP" | IMP.EXP == "IMPORT" | IMP.EXP == "EXP")

d <- d[order(d[,"VESSEL"]),]
dat <- d[!is.na(d$PORT), ]
for(a in 1:17)
{dat[,a] <- as.character(dat[,a])
dat[,a][is.na(dat[,a])] <- "Unknown" }

##Other duplicates

library(dplyr)
dati <- filter(dat,STATUS == "VESSEL AT OUTER ANCHORAGE" | STATUS == "VESSEL AT KUTUBDIA" | STATUS == "VESSEL AT ANCHOARGE" | STATUS=="VESSEL NOT ENTERING" | STATUS==" ")

##FOR SUGAR REPORTS --2=STATUS, 3=VESSEL, 6=QTY..Mts.
##FOR WHEAT REPORTS --2=STATUS, 4=VESSEL, 8=QTY..Mts.
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
de <- cleaned[colSums(!is.na(cleaned)) > 0]
##SAVING DATA TO EXCEL

library(xlsx)
write.xlsx(de,"sugar co report.xlsx",sheetName = "14,17,19-july,2016",append = TRUE)




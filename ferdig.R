rm(list=ls()); #slette alle variabler som kan være fra før
library(PxWebApiData); library(openxlsx); library(dplyr); library(tidyr)#library(jsonlite); #library(plyr);#pakke som trenges
setwd("C:/Users/e-iryku/Desktop/laste ned") #work direction 
#setwd("C:/Users/e-iryku/Trøndelag fylkeskommune/Seksjon Plan - Statistikk/R/Iryna/laste database") #work direction 


Endringen<- function(End){
    MAX<-max(sapply(End["år"], function(x) as.numeric(x) )) #Max år i tabellen
    MIN<-MAX-10      # endringen i ti år
    if(sum(End["år"]==MIN)==0) { #hvis tabell ikke inneholder 10 år, er MIN minste år
        MIN<-min(select(filter(End, år>MIN), år)[,1])
    }
    
    Endring<-End %>%
        filter(år %in% c(MAX, MIN)) %>%             #ta bare årene som trenges
        group_by_at(vars(-år, -value)) %>%   #gruppere
        summarise(endringen=diff(value), prosenet=100*round(diff(value)/sum(value), digits = 4)) %>%  #dif - differanse, prosent - differance i prosent
        as.data.frame() #riktig format av tabell
}

list_tabeller<-read.xlsx("tabellene.xlsx", colNames = TRUE, sheet = "ferdig")



for (i in 1:nrow(list_tabeller)){
    Num<-list_tabeller[[1]][i]
    print(Num)
    mappe<-pull(filter(list_tabeller, Nummer==Num), mappe)
    MetaData<-ApiData(Num, returnMetaFrames = TRUE) #lesing av metadata til tabell
    
    navn<-names(ApiData(Num, Region=1, defaultJSONquery=1, ContentsCode =1, Tid=1))[1]
    navn<-gsub(": ", "_", navn) #byte ":" med "_" fordi det er umulig å bruke ":" i navn av en mappe eller en fil
    navn<-gsub("/", ",", navn)
    
    if (is.na(list_tabeller[[3]][i])){
        TID<-TRUE
    } else{
        Tid_var<-strsplit(list_tabeller[[3]][i], split = " ")
        if (Tid_var[[1]][1] == "siste") {
            TID<--c(1:Tid_var[[1]][2])
        } else if(Tid_var[[1]][1] == "fra"){
            TID<-c(grep(Tid_var[[1]][2], MetaData$Tid[,1]):length(MetaData$Tid[,1]))
        } else if (Tid_var[[1]][1]=="til") {
            TID<-c(1:grep(Tid_var[[1]][2], MetaData$Tid[,1]))
        }
    }
    
    ikke_variabel<-list_tabeller[[4]][i]
    fjerne<-grep(list_tabeller[[5]][i], MetaData$ContentsCode$valueTexts, value=TRUE)
    
    ############################## hele landet ###############################
    
    if (sum(MetaData[[1]]=="Hele landet"|MetaData[[1]]=="Heile landet"|MetaData[[1]]=="I alt")!=0){
        if (sum(MetaData[[1]]=="Hele landet")==1){
            hel_lan<-"Hele landet"
        } else if (sum(MetaData[[1]]=="Heile landet")==1) {
            hel_lan<-"Heile landet"
        } else {
            hel_lan<-"I alt"
        }
        
        hele_landet<-ApiData(Num, 
                         Region =hel_lan,
                         defaultJSONquery=TRUE, #tatt av andre variabler som kan være "eliminate"  
                         ContentsCode=1,
                         #ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                         Tid=TID)[[1]]
    
        if (!is.na(ikke_variabel)){
            hele_landet<-hele_landet %>%
                group_by_at(vars(-c(ikke_variabel, "value"))) %>%
                summarise(value=sum(value)) %>%
                as.data.frame()
        }
        
        if (sum(!is.na(fjerne)!=0)) {
            hele_landet<-filter(hele_landet, !statistikkvariabel %in% fjerne)
        }
        
        
        #----------------------------- endringen i hele landet--------------------
        
        
        if (n_distinct(hele_landet$år)!=1){ Endringen_i_hele_landet<-Endringen(hele_landet) }
        
        
        #----------------------------- skriving på pc-en ---------------------------
        
        if (!file.exists(mappe)) {dir.create(mappe)}
        if (!file.exists(paste("./", mappe, "/", navn, sep=""))) {
            dir.create(paste("./", mappe, "/", navn, sep=""))
        } #laging av mappe med nummer av tabellen hvis det ikke finnes.
        
        write.csv(hele_landet,
                  file=paste("./",mappe, "/", navn, "/", Num, "-hele_landet.csv", sep="")) #hele landet
        write.csv(Endringen_i_hele_landet, 
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_hele_landet.csv", sep="")) #endringen
        
    }
    
    
    
    ############################## fylker #######################################
    
    fylker_koder<-c("01", "02", "03", "04", "05", "06", "07", "08", "09", "10",
                    "11", "12", "14", "15", "16", "17", "18", "19", "20", "50", "02-03", 	"19_21") #13 - Bergen, 19-21- Tromsø og Svalbard
    
    fylker_i_tabell<-intersect(fylker_koder, MetaData$Region$values)
    
    if (length(fylker_i_tabell)>0){
        fylker<-ApiData(Num, 
                        Region =fylker_i_tabell,
                        defaultJSONquery=TRUE, #tatt av andre variabler som kan være "eliminate"
                        ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                        Tid=TID)[[1]]
        
        if (!is.na(ikke_variabel)){
            fylker<-fylker %>%
                group_by_at(vars(-c(ikke_variabel, "value"))) %>%
                summarise(value=sum(value)) %>%
                as.data.frame()
        }
        
        if (sum(!is.na(fjerne))!=0) {
            fylker<-filter(fylker, !statistikkvariabel %in% fjerne)
        }
        
        
        #------------------------- Sammenslåing av Trøndelag --------------------
        
        fylker[, "region"]<-gsub("Nord-", "", fylker[, "region"]) #fjerning (bytte med ingenting) av indikasjonen av Nord-Trøndelag
        fylker[, "region"]<-gsub("Sør-", "", fylker[, "region"])
        fylker[, "region"]<-gsub(" \\(-2017\\)", "", fylker[, "region"])
        fylker[fylker[, "region"]=="Trøndelag", ]<-replace_na(fylker[fylker[, "region"]=="Trøndelag", ], list(value=0))
        
        fylker<-fylker%>%
            group_by_at(vars(-value))%>%
            summarise(value=sum(value))%>%
            as.data.frame
        
        
        #------------------------- endringen i fylker------------------------------
        
        if (n_distinct(fylker$år)!=1){ Endringen_i_fylker<-Endringen(fylker) }
        
        
        #------------------------- skriving av datasett på pc-en ---------------------
        
        if (!file.exists(mappe)) {dir.create(mappe)}
        if (!file.exists(paste("./", mappe, "/", navn, sep=""))) {
            dir.create(paste("./", mappe, "/", navn, sep=""))
        } #lagring av mappe med nummer av tabellen hvis det ikke finnes.
        write.csv(fylker, 
                  file=paste("./", mappe, "/", navn, "/",  Num, "-fylker.csv", sep=""))
        
        write.csv(filter(fylker, region=="Trøndelag"),
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Trøndelag.csv", sep=""))
        
        write.csv(Endringen_i_fylker, 
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_fylker.csv", sep=""))
        
        write.csv(filter(Endringen_i_fylker, region=="Trøndelag"),
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_Trøndelag.csv", sep=""))
    }
    
    
    ############################## kommuner i Trøndelag ################################
    
    
    if (length(grep("^50[0-9][0-9]", MetaData[[1]][,1], value=TRUE))!=0){
        if (Num!=7459 & Num!=9817 & Num!=11615){
            kommuner<-ApiData(Num, 
                              Region =list("all", c("17*", "16*", "50*", "*567")), #bare Trøndelag
                              defaultJSONquery=TRUE, #tatt av andre variable 
                              ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                              Tid=TID)[[1]]
        } else{
            kommuner1<-ApiData(Num, 
                               Region =list("all", c("17*")), #bare Trøndelag
                               defaultJSONquery=TRUE, #tatt av andre variable 
                               ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                               Tid=TID)[[1]]
            
            kommuner2<-ApiData(Num, 
                               Region =list("all", c("16*")), #bare Trøndelag
                               defaultJSONquery=TRUE, #tatt av andre variable 
                               ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                               Tid=TID)[[1]]
            kommuner3<-ApiData(Num, 
                               Region =list("all", c("50*")), #bare Trøndelag
                               defaultJSONquery=TRUE, #tatt av andre variable 
                               ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                               Tid=TID)[[1]]
            kommuner4<-ApiData(Num, 
                               Region =list("all", c("*567")), #bare Trøndelag
                               defaultJSONquery=TRUE, #tatt av andre variable 
                               ContentsCode =list("all", "*"), #tatt av alle statistikkvariabler
                               Tid=TID)[[1]]
            kommuner<-rbind(kommuner1, kommuner2, kommuner3, kommuner4)
            
        }
        
        if (!is.na(ikke_variabel)){
            kommuner<-kommuner %>%
                group_by_at(vars(-c(ikke_variabel, "value"))) %>%
                summarise(value=sum(value)) %>%
                as.data.frame()
        }
        
        if (sum(!is.na(fjerne))!=0) {
            kommuner<-filter(kommuner, !statistikkvariabel %in% fjerne)
        }
        
        
        kommuner<-kommuner[!(kommuner[, 1]=="Haltdalen (-1971)"), ]
        kommuner[, "region"]<-gsub(" \\(-2017\\)", "", kommuner[, "region"])
        kommuner<-replace_na(kommuner, list(value=0))
        
        kommuner<-kommuner%>%
            group_by_at(vars(-value))%>%
            summarise(value=sum(value))%>%
            as.data.frame
        
        Inderøy<-kommuner %>%
            filter(region %in% c("Inderøy", "Inderøy (-2011)", "Mosvik  (-2011)")) %>%
            group_by_at(vars(-region, -value)) %>%
            summarise(region="Inderøy", value=sum(value)) %>%
            as.data.frame()
        
        Indre_Fosen<-kommuner %>%
            filter(region %in% c("Indre Fosen", "Leksvik", "Rissa")) %>%
            group_by_at(vars(-region, -value)) %>%
            summarise(region="Indre Fosen", value=sum(value)) %>%
            as.data.frame()
        
        kommuner<-union(filter(kommuner, !region %in% c("Indre Fosen", "Leksvik", "Rissa", "Inderøy", "Inderøy (-2011)", "Mosvik  (-2011)")),
                        Inderøy, Indre_Fosen)
        
        #----------------------------- endringen ------------------------------------------- 
        
        if (n_distinct(kommuner$år)!=1){ Endringen_i_kommuner<-Endringen(kommuner) }
        
        #----------------------------- skriving av datasett på pc-en ------------------------
        
        if (!file.exists(mappe)) {dir.create(mappe)}
        if (!file.exists(paste("./", mappe, "/", navn, sep=""))) {
            dir.create(paste("./", mappe, "/", navn, sep=""))
        } #lagring av mappe med nummer av tabellen hvis det ikke finnes.
        
        write.csv(kommuner, 
                  file=paste("./", mappe, "/", navn, "/",  Num, "-kommuner.csv", sep=""))
        
        write.csv(Endringen_i_kommuner, 
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_fylker.csv", sep=""))
        
        
        ############################# deler av Trøndelag #####################################
        
        reg_innd<-t(read.xlsx("region inndeling.xlsx", rowNames = TRUE, colNames = FALSE)) #lesing av extern fil med regioner. resultat er data frame
        #nå ser reg_innd som data frame, med det ar bedre å ha list hvor hvert element er vektor med navn som er navn av regionen. 
        inndeling<-list()
        for (i in 1:ncol(reg_innd)){inndeling[[colnames(reg_innd)[i]]]<-as.vector(na.omit(reg_innd[,i]))} #transformasjon til list og sletting av missing elementer
        inndeling$`Trøndelag u/Trondheimsregionen`<-NULL #sletting av kolonnen med "Trondelag uten Trondheimregion"
        
        rm(reg_innd) #sletting av data frame fra lesing av extern fil
        
        Regioner<-data.frame()
        
        for (i in names(inndeling)){
            Regioner<-kommuner %>%
                filter(region %in% inndeling[[i]]) %>%
                group_by_at(vars(-region, -value)) %>%
                summarise(region=i, value=sum(value))%>%
                as.data.frame()%>%
                bind_rows(Regioner)
        }
        
        uten_Trondheimsregion<-kommuner %>%
            filter(!region %in% inndeling$Trondheimsregionen) %>%
            group_by_at(vars(-region, -value)) %>%
            summarise(region="Trøndelag uten Trondheimsregion", value=sum(value))%>%
            as.data.frame()
        
        Regioner<-bind_rows(Regioner, uten_Trondheimsregion)
        
        Regioner<-select(Regioner, names(kommuner))
        
        
        #----------------------------- endringen i regioner --------------------------------------
        
        if (n_distinct(Regioner$år)!=1){ Endringen_i_Regioner<-Endringen(Regioner) }
        
        
        #----------------------------- skriving av datasett på pc-en ------------------------------
        
        write.csv(Regioner,
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Regioner.csv", sep=""))
        
        write.csv(Endringen_i_Regioner,
                  file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_Regioner.csv", sep=""))
        
        #----------------------------- beregning av Trøndelag som summer av kommuner -------------------
        
        
        #bare for 4776
        if (sum(MetaData[[1]]=="Trøndelag")==0){
            Trøndelag<-kommuner %>%
                group_by_at(vars(-c(region, value))) %>%
                summarise(value=sum(value)) %>%
                as.data.frame() %>%
                bind_cols(data.frame(region=rep("Trøndelag", nrow(filter(kommuner, region=="Trondheim"))))) %>% 
                select(names(kommuner))
            
            if (n_distinct(Trøndelag$år)!=1){ Endringen_i_Trøndelag<-Endringen(Trøndelag) }
            
            write.csv(Trøndelag,
                      file=paste("./", mappe, "/", navn, "/",  Num, "-Trøndelag.csv", sep=""))
            
            write.csv(Endringen_i_Trøndelag,
                      file=paste("./", mappe, "/", navn, "/",  Num, "-Endringen_i_Trøndelag_.csv", sep=""))
        
        }
        
    }
}


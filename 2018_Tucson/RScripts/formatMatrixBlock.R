### Format Member Suite ###
### At this point, we've scrubbed member suite and we're dealing with the master copy off of Google Drive ###
require(xlsx)
talks2<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Templates/WorkingMaster2018TucsonSchedule_08Mar2018.xlsx",sheetName="Sheet1",stringsAsFactors=F)

load("talks.Rdata") #data.frame talks
### Genereate Short Author names or steal from other scripts
### Generate Output for Schedule Matrix Template ###
short_author_vec<-vector()
colnames(talks)

short_author_vec<-vector()
for(i in 1:nrow(talks)){
	colnames(talks)
	if(talks[i,23]=="Yes"){
		foo<-talks[rownames(talks)[i],83]
		bar<-talks[rownames(talks)[i],85:89]
		bar<-paste(bar[which(!bar=="")],collapse=", ")
		if(!bar==""){
			foo<-paste(foo,", ",bar,sep="")
			}
		}else{
			foo<-talks[rownames(talks)[i],84]
			bar<-talks[rownames(talks)[i],85:89]
			bar<-paste(bar[which(!bar=="")],collapse=", ")
			if(!bar==""){
				foo<-paste(foo,", ",bar,sep="")
			}
		}	
	short_author_vec[i]<-foo
}
names(short_author_vec)<-rownames(talks)
talks$short_author_vec<-short_author_vec

talks2$short_author_vec<-talks[talks2$Abstract_ID,]$short_author_vec

rooms<-na.omit(unique(talks2$Room))
rooms<-rooms[c(4,7,6,5,2,3,1)]

talk_list<-split(talks2,talks2$Day)
names(talk_list)

talk_list<-talk_list[2:5]

time_slots<-list(A=c("10:30","10:45","11:00","11:15","11:30","11:45"),B=c("13:30","13:45","14:00","14:15","14:30","14:45"),C=c("15:30","15:45","16:00","16:15","16:30","16:45"))
i<-1
j<-2

for(i in 1:length(talk_list)){	
	for(j in 1:length(time_slots)){
		foo<-talk_list[[i]][grep(names(time_slots)[j],talk_list[[i]]$Day.Room.TimeCode),]
		times<-gsub(":","",time_slots[[j]])
						
		sink(file=paste("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/BlockSchedule/",names(talk_list)[i],"-",names(time_slots)[j],".tex",sep=""))
		
		for(k in 1:length(times)){
			bar<-foo[foo$Time == times[k],]
			bar$Room<-factor(bar$Room,levels=rooms)		
				
			if(any(!rooms %in% bar$Room)){
				notrep<-which(!rooms %in% bar$Room)
				for(m in 1:length(notrep)){
					buck<-rep(NA,ncol(bar))
					buck[which(colnames(bar)=="Room")]<-rooms[notrep[m]]
					bar<-rbind(bar, buck)
				}
			}
			
			bar<-bar[!grepl("NA",rownames(bar)),]
						
			if(k==1){
					cat("Session &")
					cat(gsub("/","\\textbackslash ",gsub("&","\\\\&",bar$GS.Title[order(as.numeric(bar$Room))]),fixed=T), sep=" & ")
					cat("\\\\\n")
					cat("\\hline\n")
			}
			
			cat(time_slots[[j]][k],"&",sep="")
			bar$Room<-as.character(bar$Room)
			for(m in 1:length(rooms)){
				cat("\\textit{",bar[which(bar$Room==rooms[m]),]$Title,"} \\newline \\newline ", bar[which(bar$Room==rooms[m]),]$short_author_vec,sep="")
				if(m<length(rooms)){cat(" & ")}else{next}
			}
			cat("\\\\\n")
			cat("\\hline")		
		}
		sink()		
	}	
	}

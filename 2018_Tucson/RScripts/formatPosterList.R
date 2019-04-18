### Format Member Suite ###
### At this point, we've scrubbed member suite and we're dealing with the master copy off of Google Drive ###
require(xlsx)
talks2<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Templates/WorkingMaster2018TucsonSchedule.xlsx",sheetName="Sheet1",stringsAsFactors=F)

### Genereate Short Author names or steal from other scripts
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

### Create poster numbers ###
sessions<-split(talks2,talks2$Session)
poster_sessions<-sessions[grep("poster",names(sessions))]
poster_sessions<-poster_sessions[c(2,1)]
poster_sessions[[i]]$short_author_vec

head(poster_sessions[[i]])
poster_sessions[[1]]$Prez.Award

for(i in 1:length(poster_sessions)){
	sink(paste("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/PosterList/",c("wednesday","thursday")[i],".tex",sep=""))
	poster_sessions[[i]]$poster_number<-paste(i,".",poster_sessions[[i]]$Timeslot,sep="")
	for(j in 1:nrow(poster_sessions[[i]])){
		cat("\\posterentry{",poster_sessions[[i]]$poster_number[j],"}{",sep="")
		
		if(poster_sessions[[i]]$Prez.Award[j]==1){
			foo<-strsplit(poster_sessions[[i]]$short_author_vec[j]," ")[[1]]
			poster_sessions[[i]]$short_author_vec[j]<-paste(paste("*",foo[1]," ",foo[2],sep=""),foo[3],foo[4])
		}
		cat(poster_sessions[[i]]$short_author_vec[j],"}{",poster_sessions[[i]]$Title[j],"}\n",sep="")
	}
	sink()
}

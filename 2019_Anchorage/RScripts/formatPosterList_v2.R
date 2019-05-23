### Format Member Suite ###
### At this point, we've scrubbed member suite and we're dealing with the master copy off of Google Drive ###
require(xlsx)
talks<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Templates/WorkingMaster2018TucsonSchedule.xlsx",sheetName="Sheet1",stringsAsFactors=F)

### Clean Middle Names ###
midcols<-grep("middle",colnames(talks),ignore.case=T)

for(i in 1:length(midcols)){
	talks[,midcols[i]][is.na(talks[,midcols[i]])]<-""
	talks[,midcols[i]][grep("none",talks[,midcols[i]],ignore.case=T)]<-""
	talks[,midcols[i]][grep("n/a",talks[,midcols[i]],ignore.case=T)]<-""
	talks[,midcols[i]]<-gsub("[.]","",talks[,midcols[i]])
	talks[,midcols[i]]<-sapply(talks[,midcols[i]], function(x) paste(sapply(strsplit(strsplit(x," ")[[1]],""),function(x) x[1]),collapse=""))
}

### Clean First and Last Names ###
firstcols<-grep("first",colnames(talks),ignore.case=T)
lastcols<-grep("last",colnames(talks),ignore.case=T)
firstlastcols<-c(firstcols,lastcols)

for(i in 1:length(firstlastcols)){
	talks[,firstlastcols[i]][is.na(talks[,firstlastcols[i]])]<-""
	talks[,firstlastcols[i]][talks[,firstlastcols[i]]==toupper(talks[,firstlastcols[i]])]<-gsub("([[:alpha:]])([[:alpha:]]+)", "\\U\\1\\L\\2", talks[,firstlastcols[i]][talks[,firstlastcols[i]]==toupper(talks[,firstlastcols[i]])], perl=TRUE)
	talks[,firstlastcols[i]]<-gsub("[(].*?[)]","", talks[,firstlastcols[i]])
	talks[,firstlastcols[i]]<-gsub("[.]","", talks[,firstlastcols[i]])
	talks[,firstlastcols[i]] <-sapply(talks[,firstlastcols[i]], function(x) gsub("^\\s+|\\s+$", "", x))
}

colnames(talks)
### Generate Combined Author Name Field Template ###
name_vec_lists<-list()
foo<-grep("Presenting.Author.First.Name",colnames(talks))
name_vec_lists[[1]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.1...First.Name",colnames(talks))
name_vec_lists[[2]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.2...First.Name",colnames(talks))
name_vec_lists[[3]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.3...First.Name",colnames(talks))
name_vec_lists[[4]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.4...First.Name",colnames(talks))
name_vec_lists[[5]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.5...First.Name",colnames(talks))
name_vec_lists[[6]]<-c(foo,foo+1,foo+2)
foo<-grep("Co.Author.6...First.Name",colnames(talks))
name_vec_lists[[7]]<-c(foo,foo+1,foo+2)

for(i in 1:length(name_vec_lists)){
	talks[,(ncol(talks)+1)]<-rep(NA,nrow(talks))
	for(j in 1:nrow(talks)){
		talks[j,ncol(talks)]<-gsub("\\s+"," ",paste(talks[j,name_vec_lists[[i]]],collapse=" "))
	}
}

colnames(talks)[(ncol(talks)-6):ncol(talks)]<-c("Presenting Author Full Name",paste("Author",1:6,"Full Name"))
head(talks)

### Create SHORT version of Author Names for Block Schedules ###
for(i in 1:length(name_vec_lists)){
	talks[,(ncol(talks)+1)]<-rep("",nrow(talks))
	
	for(j in 1:nrow(talks)){
		if(talks[j,name_vec_lists[[i]][3]]==""){
			next
			}else{
				if(talks[j,name_vec_lists[[i]][2]]==""){
					talks[j,ncol(talks)]<-paste(talks[j,name_vec_lists[[i]][3]]," ",strsplit(talks[j,name_vec_lists[[i]][1]],"")[[1]][1],sep="")
					}else{
						talks[j,ncol(talks)]<-paste(talks[j,name_vec_lists[[i]][3]]," ",strsplit(talks[j,name_vec_lists[[i]][1]],"")[[1]][1],strsplit(talks[j,name_vec_lists[[i]][2]],"")[[1]][1],sep="")
						}
					}
				}
			}

colnames(talks)[(ncol(talks)-6):ncol(talks)]<-c("Presenting Author Short Name",paste("Author",1:6,"Short Name"))
colnames(talks)

### Create combined vectors of short and long author names ###
#talks<-talks[,-((ncol(talks)-1):ncol(talks))]
talks[,ncol(talks)+1]<-rep(NA,nrow(talks))
talks[,ncol(talks)+1]<-rep(NA,nrow(talks))

fullnamecols<-grep("Full Name",colnames(talks))
shortnamecols<-grep("Short Name",colnames(talks))
i<-8
for(i in 1:nrow(talks)){
	if(talks$"Checkbox...Same.as.Presenting.Author"[i]=="Yes"){
		emptycols<-which(talks[i,c(fullnamecols[-2])] == " " | talks[i, fullnamecols[-2]] == "")
		
		if(length(emptycols)>=1){
			talks[i,(ncol(talks)-1)]<-paste(talks[i, fullnamecols[-2][-emptycols]],collapse=", ")			
			talks[i,(ncol(talks))]<-paste(talks[i, shortnamecols[-2][-emptycols]],collapse=", ")			
			}else{
				talks[i,(ncol(talks)-1)]<-paste(talks[i, fullnamecols[-2]],collapse=", ")			
				talks[i,(ncol(talks))]<-paste(talks[i, shortnamecols[-2]],collapse=", ")			
			}
		}else{
			emptycols<-which(talks[i,fullnamecols] == " " | talks[i,fullnamecols] == "")
	
			if(length(emptycols)>=1){
				talks[i,(ncol(talks)-1)]<-paste(talks[i, fullnamecols[-emptycols]],collapse=", ")			
				talks[i,(ncol(talks))]<-paste(talks[i, shortnamecols[-emptycols]],collapse=", ")			
			}else{
					talks[i,(ncol(talks)-1)]<-paste(talks[i, fullnamecols],collapse=", ")			
					talks[i,(ncol(talks))]<-paste(talks[i, shortnamecols],collapse=", ")			
				}
			}	
		}
				
colnames(talks)[(ncol(talks)-1):ncol(talks)]<-c("FullLongAuthor","FullShortAuthor")
head(talks,20)

talks$FullShortAuthor

###
talks$FullLongAuthor <-gsub("á","\\\\'{a}",talks$FullLongAuthor)
talks$FullShortAuthor<-gsub("á","\\\\'{a}",talks$FullShortAuthor)
talks$FullLongAuthor <-gsub("ó","\\\\'{o}",talks$FullLongAuthor)
talks$FullShortAuthor<-gsub("ó","\\\\'{o}",talks$FullShortAuthor)
talks$FullLongAuthor <-gsub("í","\\\\'{i}",talks$FullLongAuthor)
talks$FullShortAuthor<-gsub("í","\\\\'{i}",talks$FullShortAuthor)
talks$FullLongAuthor <-gsub("é","\\\\'{e}",talks$FullLongAuthor)
talks$FullShortAuthor<-gsub("é","\\\\'{e}",talks$FullShortAuthor)
talks$FullLongAuthor <-gsub("ñ","\\\\~{n}",talks$FullLongAuthor)
talks$FullShortAuthor<-gsub("ñ","\\\\~{n}",talks$FullShortAuthor)

### Deal with comma issues ###
for(i in 1:nrow(talks)){
	foo<-strsplit(talks$FullLongAuthor[i],"")[[1]]
	if(length(foo) == 0){next}
	
	if(foo[length(foo)-1]==","){
		foo<-foo[1:(length(foo)-2)]
	}
	
	foo<-paste(foo,collapse="")
	talks$FullLongAuthor[i]<-foo
	
	bar<-strsplit(talks$FullShortAuthor[i],"")[[1]]
	if(bar[length(bar)-1]==","){
		bar<-bar[1:(length(bar)-2)]
	}
	bar<-paste(bar,collapse="")
	talks$FullShortAuthor[i]<-bar
}

strsplit(talks$FullLongAuthor,"")[[1]]
talks$FullShortAuthor

### Remove retracted talks ###
talks<-talks[!talks$Status == "Retracted",]

### Create poster numbers ###
sessions<-split(talks,talks$Session)
poster_sessions<-sessions[grep("poster",names(sessions))]
poster_sessions<-poster_sessions[c(2,1)]


i<-1
j<-41

for(i in 1:length(poster_sessions)){
	
	sink(paste("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/PosterList/",c("wednesday","thursday")[i],".tex",sep=""))
	poster_sessions[[i]]$poster_number<-paste(i,".",poster_sessions[[i]]$Timeslot,sep="")
	poster_sessions[[i]]<-poster_sessions[[i]][order(as.numeric(poster_sessions[[i]]$Timeslot)),]
	
	for(j in 1:nrow(poster_sessions[[i]])){
		if(poster_sessions[[i]]$paperPoster[j]=="1"){
			cat("\\posterentry{",poster_sessions[[i]]$poster_number[j],"}{",sep="")
	
			if(poster_sessions[[i]]$Prez.Award[j]==1){
				poster_sessions[[i]]$FullShortAuthor[j]<-paste("*", poster_sessions[[i]]$FullShortAuthor[j],sep="")
			}
			
			cat(poster_sessions[[i]]$FullShortAuthor[j],"}{",poster_sessions[[i]]$Title[j],"}\n",sep="")
		}else{
			next}
		}	
	sink()
}

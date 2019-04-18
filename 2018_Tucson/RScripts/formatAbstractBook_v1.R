### Format Member Suite ###
### At this point, we've scrubbed member suite and we're dealing with the master copy off of Google Drive ###
require(xlsx)
require(Hmisc)

talks<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Templates/WorkingMaster2018TucsonSchedule_24March for Abstract Book.xlsx",sheetName="Sheet1",stringsAsFactors=F)

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
talks$FullLongAuthor
talks$FullShortAuthor

### Remove retracted talks ###
talks<-talks[!talks$Status == "Retracted",]

### Clean up abstracts ###
#gsub(talks$Abstract)

### Split talks into separate data.frames ###
posters<-talks[talks$GS.Title=="Poster",]

### Sort by last name of presenting author ###
posters<-posters[order(posters$Presenting.Author.Last.Name),]
posters$FullLongAuthor<-gsub(", ","\\\\\\\\",posters$FullLongAuthor)

firstafoo<-sapply(strsplit(sapply(strsplit(posters$FullLongAuthor,"\\\\\\\\"),function(x) x[1])," "),function(x) x[length(x)])
posters<-posters[order(firstafoo),]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/posters.tex")
for(i in 1:nrow(posters)){
	cat("\\normaltalk")
	cat("{",posters$Title[i],"}",sep="")
	cat("{",posters$FullLongAuthor[i],"}",sep="")
	cat("{",posters$Abstract[i],"}",sep="")
	cat("\n\n")
}

sink()

### Lightning Talks ###
l_talks<-talks[talks$GS.Title=="Lightning Talks",]

### Sort by last name of presenting author ###
l_talks <-l_talks[order(l_talks $Presenting.Author.Last.Name),]
l_talks $FullLongAuthor<-gsub(", ","\\\\\\\\", l_talks $FullLongAuthor)

firstafoo <-sapply(strsplit(sapply(strsplit(l_talks $FullLongAuthor,"\\\\\\\\"),function(x) x[1])," "),function(x) x[length(x)])
l_talks <-l_talks[order(firstafoo),]

l_talks<-l_talks[!l_talks[,1]=="VACANT",]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/lightningtalks.tex")
for(i in 1:nrow(l_talks)){
	cat("\\normaltalk")
	cat("{", l_talks $Title[i],"}",sep="")
	cat("{", l_talks $FullLongAuthor[i],"}",sep="")
	cat("{", l_talks $Abstract[i],"}",sep="")
	cat("\n\n")
}

sink()

### "NORMAL" Talks ###
n_talks<-talks[!(talks$GS.Title=="Lightning Talks" | talks$GS.Title=="Poster" |  talks$GS.Title==""),]
n_talks$GS.Title

### Sort by last name of presenting author ###
n_talks <-n_talks[order(n_talks $Presenting.Author.Last.Name),]
n_talks $FullLongAuthor<-gsub(", ","\\\\\\\\", n_talks $FullLongAuthor)

firstafoo <-sapply(strsplit(sapply(strsplit(n_talks $FullLongAuthor,"\\\\\\\\"),function(x) x[1])," "),function(x) x[length(x)])
n_talks <-n_talks[order(firstafoo),]

n_talks <-n_talks[! n_talks[,1]=="VACANT",]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/normaltalks.tex")
for(i in 1:nrow(n_talks)){
	cat("\\normaltalk")
	cat("{", n_talks $Title[i],"}",sep="")
	cat("{", n_talks $FullLongAuthor[i],"}",sep="")
	cat("{", n_talks $Abstract[i],"}",sep="")
	cat("\n\n")
}

sink()
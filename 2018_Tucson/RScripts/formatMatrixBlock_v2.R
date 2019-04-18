### Format Member Suite ###
### At this point, we've scrubbed member suite and we're dealing with the master copy off of Google Drive ###
require(xlsx)
require(Hmisc)

talks<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Templates/WorkingMaster2018TucsonSchedule_24March.xlsx",sheetName="Sheet1",stringsAsFactors=F)

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

rooms<-c("Agave II-III","Coronado I","Coronado II","Presidio I","Presidio II","Presidio III-IV","Presidio V")
rooms<-factor(rooms,levels=rooms)

talk_list<-split(talks,talks$Day)
talk_list<-talk_list[-which(names(talk_list)=="")]
names(talk_list)

time_slots<-list(A=c("10:30","10:45","11:00","11:15","11:30","11:45"),B=c("13:30","13:45","14:00","14:15","14:30","14:45"),C=c("15:30","15:45","16:00","16:15","16:30","16:45"))

### Create tex files 
i<-3
j<-1
k

for(i in 1:length(talk_list)){	
	for(j in 1:length(time_slots)){
		foo<-talk_list[[i]][grep(names(time_slots)[j],talk_list[[i]]$Day.Room.TimeCode),]
		times<-gsub(":","",time_slots[[j]])
		
		sink(file=paste("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/BlockSchedule/",names(talk_list)[i],"-",names(time_slots)[j],".tex",sep=""))
		
		spec_col<-which((rooms) %in% unique(foo[grep("Sym",foo$GS.Title),]$Room))
		
		### Create tabular string with special column types for symposia ###
		col_color_vec<-rep(1,length(rooms))
		col_color_vec[spec_col]<-2 #These are the columns that require special color		
		
		cat("\\begin{tabular}{|x{0.85cm}")
		for(k in 1:length(col_color_vec)){
			cat("|")
			cat(c("x","a")[col_color_vec[k]])
			cat("{2.85cm}")
		}
		cat("|}\\hline\n")

		cat(c("Room",as.character(rooms)),sep=" & ")
		cat("\\\\\n")
		cat("\\hline\n")
		
		if("Sym: Early Professional Mini Talks" %in% foo$GS.Title){
			trix<-foo[which(foo$GS.Title == "Sym: Early Professional Mini Talks"),]
			trix_times<-c("1530","1545","1600","1615","1630","1645")
			trix$fullauthor<-paste(trix$Presenting.Author.First.Name,trix$Presenting.Author.Middle.Name,trix$Presenting.Author.Last.Name,sep=" ")
			
			for(k in 1:length(trix_times)){
				rec<-(k-1)*3+1
				rec<-(rec:(rec+2))
				
				times_form<-strsplit(trix$Time[order(as.numeric(trix$Timeslot))][rec],"")
				times_form<-sapply(times_form,function(x) paste(c(x[1:2],":",x[3:4]),collapse=""))

				trix_title<-paste(trix$fullauthor[order(as.numeric(trix$Timeslot))][rec]," (", times_form ,")",sep="",collapse="\\newline \\newline \\newline ")
				trix_title<-gsub("(NA)","",trix_title)
				trix_title<-gsub("&","\\\\&",trix_title)
				trix_title<-gsub("(16:50)","",trix_title,fixed=T)
				trix_title<-gsub("(16:55)","",trix_title,fixed=T)
				
				foo[which(foo$GS.Title == "Sym: Early Professional Mini Talks" & foo$Time==trix_times[k]),]$Title<-trix_title
			}
		}
		
		if("Lightning Talks" %in% foo$GS.Title){
			trix<-foo[which(foo$GS.Title == "Lightning Talks"),]
			trix_times<-c("1030","1045","1100","1115","1130","1145")
			
			trix$FullShortAuthor<-trix$"Presenting Author Short Name"
			
			for(k in 1:nrow(trix)){
				nempty<-length(which(as.character(trix[k,grep("Full Name",colnames(trix))])==" "))
				if(nempty<6){
					trix$"FullShortAuthor"[k]<-paste(trix$"FullShortAuthor"[k]," et al.")
				}
				if(trix$Prez.Award[k]=="1"){
					trix$"FullShortAuthor"[k]<-paste("*",trix$"FullShortAuthor"[k],sep="")
				}
			}
			
			for(k in 1:length(trix_times)){
				rec<-(k-1)*3+1
				rec<-(rec:(rec+2))
				
				times_form<-strsplit(trix$Time[order(as.numeric(trix$Timeslot))][rec],"")
				times_form<-sapply(times_form,function(x) paste(c(x[1:2],":",x[3:4]),collapse=""))
				
				trix_title<-paste(trix$FullShortAuthor[order(as.numeric(trix$Timeslot))][rec], "\\newline \\tiny ",trix$Short.Title[order(as.numeric(trix$Timeslot))][rec]," \\scriptsize",sep="",collapse="\\newline \\newline ")
				trix_title<-gsub("NA","",trix_title)
				trix_title<-gsub("&","\\\\&",trix_title)
				trix_title<-gsub("(10:30)","",trix_title,fixed=T)
				trix_title<-gsub("(11:45)","",trix_title,fixed=T)
				trix_title<-gsub("(11:50)","",trix_title,fixed=T)
				trix_title<-gsub("(11:55)","",trix_title,fixed=T)
				trix_title<-gsub("(:)","",trix_title,fixed=T)
				

				foo[which(foo$GS.Title=="Lightning Talks"),]
				foo[which(foo$GS.Title == "Lightning Talks" & foo$Time==trix_times[k]),]$Title<-trix_title
				foo[which(foo$GS.Title == "Lightning Talks" & foo$Time==trix_times[k]),]$FullShortAuthor<-""

			}
		}
	
		for(k in 1:length(times)){
			bar<-foo[foo$Time == times[k],]
			bar$Room<-factor(bar$Room,levels=rooms)		
			
			if(any(!rooms %in% bar$Room)){
				notrep<-which(!rooms %in% bar$Room)
				for(m in 1:length(notrep)){
					buck<-rep(NA,ncol(bar))
					buck[which(colnames(bar)=="Room")]<-as.character(rooms)[notrep[m]]
					bar<-rbind(bar, buck)
				}
			}
			
			bar<-bar[!grepl("NA",rownames(bar)),]
			bar<-bar[order(as.numeric(bar$Room)),]
			
			bar$GS.Title[is.na(bar$GS.Title)]<-""
			
			if(k==1){
					cat("\\rule{0pt}{1em} ")			
					cat("\\textbf{Session} &")
					symptite<-gsub("/","\\textbackslash ",gsub("&","\\\\&",bar$GS.Title),fixed=T)
					symptite<-paste("\\footnotesize \\textbf{",symptite,"}",sep="")
					cat(symptite, sep=" & ")
					cat("\\\\\n")
					#cat("\\rule{0pt}{1em} ")			
					cat("\\hline\n")
			}
			
			### 
			cat("\\makecell{",time_slots[[j]][k],"}&",sep="")
			
			bar$Room<-as.character(bar$Room)
			
			### Check if first author is competing for Award Talk ###
			### Also check for 30 min talk and create multi col & /cline if so ###
			
			bar$X30.min.talk.[is.na(bar$X30.min.talk.)]<-""
			cline_vec<-vector()
			
			for(m in 1:length(rooms)){
				if(!is.na(bar$Prez.Award[m])){
					if(bar$Prez.Award[m]=="1"){
						bar$Title[m]<-paste("*",bar$Title[m],sep="")
						bar$Title[m]<-gsub("**","*",bar$Title[m],fixed=T)
					}
				}	
				bar$Title[is.na(bar$Title)]<-""
				
				title_foo<-bar[which(bar$Room==rooms[m]),]$Title
				author_foo<-bar[which(bar$Room==rooms[m]),]$FullShortAuthor
				
				if(identical(title_foo,character(0))){
					title_foo<-""
				}
				
				if(is.na(title_foo)){
					title_foo<-""
				}
				
				if(identical(author_foo,character(0))){
					title_foo<-""
					author_foo<-""
				}
				
				if(is.na(author_foo)){
					try(author_foo <-"")
				}
				
				if(title_foo==""){
					author_foo<-""
				}
				
				if(grepl("Discussion",author_foo)){
					author_foo<-""
				}
				
				if("Sym: Early Professional Mini Talks" %in% bar$GS.Title[m]){
					author_foo<-""
				}
				
				if(bar$GS.Title[m] == "Lightning Talks"){
				title_author_foo<-paste(title_foo,sep="")
				}else{
				title_author_foo<-paste(title_foo," \\newline \\newline ", "\\textit{",author_foo,"}",sep="")
				}
				
				if(bar$X30.min.talk.[m]=="1"){
					title_author_foo<-paste("\\multirow{2}{2.95cm}{",title_author_foo,"}",sep="")
					cline_vec<-c(cline_vec,m)
				}
				cat(title_author_foo)
				if(m<length(rooms)){cat(" & ")}else{next}
			}
			
			cat("\\\\\n")
			
			### Write out \hhline or \hline ###
			
			if("Sym: Early Professional Mini Talks" %in% bar$GS.Title & k < length(times)){
				cline_vec<-c(cline_vec,6)
				}	
				
			if("Lightning Talks" %in% bar$GS.Title & k < length(times)){
				hhline_start<-"\\hhline{|->{\\arrayrulecolor{white}}|->{\\arrayrulecolor{black}}"
					}else{
					hhline_start <-"\\hhline{|-|-"
						}
					
				if(length(cline_vec)>=1){
					rle_foo<-rle(!(3:8) %in% (cline_vec+1))
					cat(hhline_start)
					for(m in 1:length(rle_foo$lengths)){
						if(rle_foo$values[m]){
							cat(paste("*{",rle_foo$lengths[m],"}",sep=""),"{|-}",sep="")
						}else{
							cat(rep(paste(">{\\arrayrulecolor{lightpurple}}|->{\\arrayrulecolor{black}}"),times=rle_foo$lengths[m]))
						}
					}
					cat("|}\n")
					
					}else{
						if("Lightning Talks" %in% bar$GS.Title & k < length(times)){
							cat(hhline_start,"|-|-|-|-|-|-}",sep="")
							}else{
					cat("\\hline\n")
					}
				}
			}
		
		if(j==1){
			cat("\\multicolumn{8}{|c|}{\\small LUNCH BREAK}\\\\\n\n")
			cat("\\hline\n")
		}
		if(j==2){
			cat("\\multicolumn{8}{|c|}{\\small COFFEE BREAK}\\\\\n\n")
				cat("\\hline\n")
		}
		cat("\\end{tabular}\n")
		sink()		
	}	
}

i
j
k
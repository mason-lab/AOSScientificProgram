### Clean up excel file from MemberSuite for LaTeX and upload to Sched.doc ###
require(xlsx)
require(XML)
require(tools)
require(Hmisc)
~
### Read in xlsx ###
talks<-read.xlsx2("~/AOSScientificProgram/2019_Anchorage/AOS 2019 Abstracts_Master_24 May 2019.v3.xlsx",sheetName="Search Results",stringsAsFactors=F)

head(talks)
colnames(talks)

### Remove weird characters from Abstracts ###
talks$Abstract<-gsub("","---",talks$Abstract)

### Remove weird characters from Title ###
talks$Title<-gsub("\\n"," ",talks$Title)
talks$Title<-gsub("&","\\\\&",talks$Title)

for(i in 1:length(talks$Title)){
	
	### Skip entries with no title ###
	if(talks$Title[i]==""){next}
	
	if(grepl("\\.",strsplit(talks$Title[i],"")[[1]][length(strsplit(talks$Title[i],"")[[1]])])){ #Removes periods at end
		talks$Title[i]<-paste(strsplit(talks$Title[i],"")[[1]][1:(length(strsplit(talks$Title[i],"")[[1]])-1)],collapse="")
	}
	if(grepl(" ",strsplit(talks$Title[i],"")[[1]][length(strsplit(talks$Title[i],"")[[1]])])){
		talks$Title[i]<-paste(strsplit(talks$Title[i],"")[[1]][1:(length(strsplit(talks$Title[i],"")[[1]])-1)],collapse="")
	}
}	

talks$Title[talks$Title==toupper(talks$Title)]<-toTitleCase(gsub("([[:alpha:]])([[:alpha:]]+)", "\\U\\1\\L\\2", talks$Title[talks$Title==toupper(talks$Title)], perl=TRUE))

### Clean Middle Names ###
midcols<-grep("middle",colnames(talks),ignore.case=T)

for(i in 1:length(midcols)){
	talks[,midcols[i]][is.na(talks[,midcols[i]])]<-""
	talks[,midcols[i]][grep("none",talks[,midcols[i]],ignore.case=T)]<-""
	talks[,midcols[i]][grep("n/a",talks[,midcols[i]],ignore.case=T)]<-""
	talks[,midcols[i]]<-gsub("[.]","",talks[,midcols[i]])
	talks[,midcols[i]]<-gsub("-","",talks[,midcols[i]])
	talks[,midcols[i]]<-gsub("_","",talks[,midcols[i]])
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

### Create Cleaned Author Strings [both FULL and SHORT versions] ###
name_vec_lists<-lapply(seq(55,121,6),function(x) (x:(x+2)))

for(i in 1:length(name_vec_lists)){
	talks[,(ncol(talks)+1)]<-rep(NA,nrow(talks))
	for(j in 1:nrow(talks)){
		talks[j,ncol(talks)]<-gsub("\\s+"," ",paste(talks[j,name_vec_lists[[i]]],collapse=" "))
	}
}

colnames(talks)[(ncol(talks)-11):ncol(talks)]<-paste("Author",1:12,"Full Name")
talks[,(ncol(talks)-11):ncol(talks)]

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

colnames(talks)[(ncol(talks)-11):ncol(talks)]<-paste("Author",1:12,"Short Name")

### Create combined vectors of short and long author names ###
talks[,ncol(talks)+1]<-rep(NA,nrow(talks))
talks[,ncol(talks)+1]<-rep(NA,nrow(talks))

fullnamecols<-grep("Full Name",colnames(talks))
shortnamecols<-grep("Short Name",colnames(talks))

for(i in 1:nrow(talks)){
	talks[i,(ncol(talks)-1)]<-paste(talks[i,fullnamecols][!as.character(talks[i,fullnamecols])==" "],collapse=", ")
	talks[i,ncol(talks)]<-paste(talks[i,shortnamecols][!as.character(talks[i, shortnamecols])==""],collapse=", ")
}	
colnames(talks)[(ncol(talks)-1):ncol(talks)]<-c("FullLongAuthor","FullShortAuthor")

### Fix accents ###
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

### Get Room Info for Talks ###
### RETURN TO THIS ONCE ROOMS HAVE BEEN ASSIGNED AND WE WANT to reordeR COLUMNS ACCORDINGLy ###
talks$Room<-factor(talks$Room,levels=1:12)

### Create tex files for talk matrix ###
posters<-talks[talks$FORMAT=="Poster",]
talks<-talks[!(talks$FORMAT=="Poster"),]

### Set up three-tiered for loop to create input tex files for each day, roomset, and timeslot ###
talks$RoomSet<-rep(NA,nrow(talks))
talks$RoomSet[talks$Room %in% 1:6]<-"RoomSetA"
talks$RoomSet[talks$Room %in% 7:12]<-"RoomSetB"

day_list<-split(talks,talks$Day)
#day_list<-day_list[-which(names(day_list)=="")]
day_room_list<-lapply(day_list,function(x) split(x,x$RoomSet))

length(day_list)
length(day_room_list[[3]])

### Set up Time Slots
time_slots<-list(TimeSetA=c("10:30","10:45","11:00","11:15","11:30","11:45"),TimeSetB=c("14:00","14:15","14:30","14:45","15:00","15:15"),TimeSetC=c("16:00","16:15","16:30","16:45","17:00","17:15"))

### Create tex files for matrix, 2 per day (1 for rooms 1:6, another for rooms 7:12) ###
i<-3
j<-2
k<-1

for(i in 1:length(day_room_list)){
	for(j in 1:length(day_room_list[[i]])){
		for(k in 1:length(time_slots)){
			sink(file=paste("~/Desktop/DesktopClutter25Oct/Service/AOS/AOS_ScientificProgramCommittee/2019/TeX/",names(day_room_list)[i],"-",names(day_room_list[[i]])[j],"-",names(time_slots)[k],".tex",sep=""))
			
			times<-gsub(":","",time_slots[[k]])
			
			this_page<-day_room_list[[i]][[j]][day_room_list[[i]][[j]]$Time %in% times,]
			this_page$Room<-factor(this_page$Room,levels=min(as.numeric(as.character(this_page$Room))):max(as.numeric(as.character(this_page$Room))))
			
			### Ascertain which columns are symposia and color columns accordingly ###
			symp_rooms<-as.numeric(unique(this_page$Room[this_page$FORMAT=="Symposium"]))
			col_color_vec<-rep(1,length(levels(this_page$Room)))
			col_color_vec[symp_rooms]<-2
									
			cat("\\begin{tabular}{|x{0.75cm}")
			for(m in 1:length(col_color_vec)){
				cat("|")
				cat(c("x","a")[col_color_vec[m]])
				cat("{2.25cm}")
			}
			cat("|}\\hline\n")
			
			cat(c("Room",levels(this_page$Room)),sep=" & ")
			cat("\\\\\n")
			cat("\\hline\n")
			
			### Deal with EP Mini Symposium Later ### 22 May 2019
			
			### Format LIGHTNING talks if present ###
			if(any(this_page$Session.Title == "Lightning Talks")){
				light_time_slots<-unlist(lapply(as.numeric(unique(this_page$Time)),function(x) seq(x,x+10,5)))
				this_page<-day_room_list[[i]][[j]][day_room_list[[i]][[j]]$Time %in% light_time_slots,]
				
				light_titles<-paste0("\\scriptsize ",this_page[this_page$Session.Title == "Lightning Talks",]$"Author 1 Short Name"," et al. \\newline \\tiny ", this_page[this_page$Session.Title == "Lightning Talks",]$"LightningTitle")
				
				light_titles_full<-sapply(lapply(seq(1,13,3),function(x) x:(x+2)),function(x) paste(light_titles[x],collapse="\\newline \\newline "))
				this_page[this_page$Session.Title == "Lightning Talks",]$Title[seq(1,13,3)]<-light_titles_full
				}
			this_page<-this_page[this_page$Time %in% times,]
			this_page$Room<-factor(this_page$Room,levels=min(as.numeric(as.character(this_page$Room))):max(as.numeric(as.character(this_page$Room))))
			
			for(m in 1:length(times)){
				this_time<-this_page[this_page$Time == times[m],]
				
				this_time$Room<-factor(this_time$Room,levels=(min(as.numeric(as.character(this_time$Room))):max(as.numeric(as.character(this_time$Room)))))
				
				this_time<-this_time[order(as.numeric(this_time$Room)),]
				
				### Make session header for this page ###	
				if(m==1){
						cat("\\rule{0pt}{1em} ")			
						cat("\\textbf{Session} &")
						symptite<-gsub("/","\\textbackslash ",gsub("&","\\\\&", this_time $Session.Title),fixed=T)
						symptite<-paste("\\footnotesize \\textbf{",symptite,"}",sep="")
						cat(symptite, sep=" & ")
						cat("\\\\\n")
						#cat("\\rule{0pt}{1em} ")			
						cat("\\hline\n")
				}
				
				### Write out time in first column ### 
				cat("\\makecell{",time_slots[[k]][m],"}&",sep="")
							
				### Check if first author is competing for Award Talk ###
				if(any(this_time$Student.Prez.Award.Competitors.1=="1")){
					this_time$Title[this_time$Student.Prez.Award.Competitors.1=="1"]<-paste0("*",this_time$Title[this_time$Student.Prez.Award.Competitors.1=="1"])
				}
				
				### Also check for 30 min talk and create multi col & /cline if so ###
				cline_vec<-vector()
				for(n in 1:nrow(this_time)){
					if(this_time$Session.Title[n]=="Lightning Talks"){
						cat(this_time$Title[n])
					}else{
						cat(this_time$Title[n]," \\newline \\newline ", "\\textit{", this_time$FullShortAuthor[n],"}",sep="")
						}
					if(n<nrow(this_time)){cat(" & ")}else{next}
				}
			cat("\\\\\n\\hline\n")
			}
			
			
			if(k==1){
				cat("\\multicolumn{7}{|c|}{\\small LUNCH BREAK}\\\\\n\n")
				cat("\\hline\n")
			}
			if(k==2){
				cat("\\multicolumn{7}{|c|}{\\small COFFEE BREAK}\\\\\n\n")
					cat("\\hline\n")
			}
			cat("\\end{tabular}\n")
			sink()			
		}
	}
}
	
# ### Split output into talks / posters ###


# ### Create table with relevant outputs to manually clean ###
# colnames(talks)

# talks_cleaned<-talks[c(143:167,)]

# write.xlsx(talks_cleaned, "/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_Submissions_Cleaned_01Mar2018_NAM.xlsx")

# ### Generate inputs for matrix schedule ###



# ### Read in Cleaned XLSX ###
# talks_man<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/CleanedMSOutput/Tucson2018_Submissions_Cleaned_01Mar2018_NAM.xlsx",sheetName="Sheet1",row.names=1,stringsAsFactors=F)
# talks_man[is.na(talks_man)]<-""
# talks_man[talks_man==" "]<-""

# aff_talks_man<-grep("Affiliation",colnames(talks_man))

# for(i in 1:length(aff_talks_man)){
	# talks_man[,aff_talks_man[i]]<-gsub("&","\\\\&",talks_man[,aff_talks_man[i]])
# }

# talks_man$Abstract<-gsub("’","'",talks_man$Abstract)
# talks_man$Abstract<-gsub("~","{\\raise.17ex\\hbox{$\\scriptstyle\\mathtt{\\sim}$}}",talks_man$Abstract)
# talks_man$Title<-gsub("’","'",talks_man$Title)
# talks_man$Title<-gsub("&","\\\\&", talks_man$Title)

# talks_man$Abstract<-latexTranslate(talks_man$Abstract)

# ### Split into different talk types ###
# talks_list<-split(talks_man,talks_man$submissionType)

# ### NORMAL TALKS ###
# ### Order according to last name of presenting author ###
# talks_list[[1]]<-talks_list[[1]][order(talks[rownames(talks_list[[1]]),6]),]

# sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/oralpresentations.tex")
# for(i in 1:nrow(talks_list[[1]])){
	# cat("\\normaltalk")
	# if(talks_list[[1]][i,17]=="Yes"){
		# cat("{",talks_list[[1]][i,3],", ",talks_list[[1]][i,4],"}",sep="")
	# }else{
		# cat("{",talks_list[[1]][i,5],", ",talks_list[[1]][i,6],"}",sep="")
	# }
	# cat("{",talks_list[[1]][i,7],", ",talks_list[[1]][i,8],"}",sep="")
	# cat("{",talks_list[[1]][i,9],", ",talks_list[[1]][i,10],"}",sep="")
	# cat("{",talks_list[[1]][i,11],", ",talks_list[[1]][i,12],"}",sep="")
	# cat("{",talks_list[[1]][i,13],", ",talks_list[[1]][i,14],"}",sep="")
	# cat("{",talks_list[[1]][i,15],", ",talks_list[[1]][i,16],"}",sep="")

	# cat("{",talks_list[[1]][i,20],"}",sep="")			
	# cat("{",talks_list[[1]][i,21]," [",rownames(talks_list[[1]])[i],"]}",sep="")	
			
	# cat("\n\n")
# }

# sink()

# ### LIGHTNING TALKS ###
# ### Order according to last name of presenting author ###
# talks_list[[2]]<-talks_list[[2]][order(talks[rownames(talks_list[[2]]),6]),]

# sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/lightningtalks.tex")
# for(i in 1:nrow(talks_list[[2]])){
	# cat("\\normaltalk")
	# if(talks_list[[2]][i,17]=="Yes"){
		# cat("{",talks_list[[2]][i,3],", ",talks_list[[2]][i,4],"}",sep="")
	# }else{
		# cat("{",talks_list[[2]][i,5],", ",talks_list[[2]][i,6],"}",sep="")
	# }
	# cat("{",talks_list[[2]][i,7],", ",talks_list[[2]][i,8],"}",sep="")
	# cat("{",talks_list[[2]][i,9],", ",talks_list[[2]][i,10],"}",sep="")
	# cat("{",talks_list[[2]][i,11],", ",talks_list[[2]][i,12],"}",sep="")
	# cat("{",talks_list[[2]][i,13],", ",talks_list[[2]][i,14],"}",sep="")
	# cat("{",talks_list[[2]][i,15],", ",talks_list[[2]][i,16],"}",sep="")

	# cat("{",talks_list[[2]][i,20],"}",sep="")			
	# cat("{",talks_list[[2]][i,21]," [",rownames(talks_list[[2]])[i],"]}",sep="")	
			
	# cat("\n\n")
# }

# sink()

# ### POSTERS ###
# ### Order according to last name of presenting author ###
# talks_list[[3]]<-talks_list[[3]][order(talks[rownames(talks_list[[3]]),6]),]

# sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/posters.tex")
# for(i in 1:nrow(talks_list[[3]])){
	# cat("\\normaltalk")
	# if(talks_list[[3]][i,17]=="Yes"){
		# cat("{",talks_list[[3]][i,3],", ",talks_list[[3]][i,4],"}",sep="")
	# }else{
		# cat("{",talks_list[[3]][i,5],", ",talks_list[[3]][i,6],"}",sep="")
	# }
	# cat("{",talks_list[[3]][i,7],", ",talks_list[[3]][i,8],"}",sep="")
	# cat("{",talks_list[[3]][i,9],", ",talks_list[[3]][i,10],"}",sep="")
	# cat("{",talks_list[[3]][i,11],", ",talks_list[[3]][i,12],"}",sep="")
	# cat("{",talks_list[[3]][i,13],", ",talks_list[[3]][i,14],"}",sep="")
	# cat("{",talks_list[[3]][i,15],", ",talks_list[[3]][i,16],"}",sep="")

	# cat("{",talks_list[[3]][i,20],"}",sep="")			
	# cat("{",talks_list[[3]][i,21]," [",rownames(talks_list[[3]])[i],"]}",sep="")	
			
	# cat("\n\n")
# }

# sink()

# ### Generate Output for Schedule Matrix Template ###
# short_author_vec<-vector()
# colnames(talks)

# for(i in 1:nrow(talks_list[[1]])){
	# colnames(talks_list[[1]])
	# if(talks_list[[1]][i,17]=="Yes"){
		# foo<-talks[rownames(talks_list[[1]])[i],83]
		# bar<-talks[rownames(talks_list[[1]])[i],85:89]
		# bar<-paste(bar[which(!bar=="")],collapse=", ")
		# if(!bar==""){
			# foo<-paste(foo,", ",bar,sep="")
			# }
		# }else{
			# foo<-talks[rownames(talks_list[[1]])[i],84]
			# bar<-talks[rownames(talks_list[[1]])[i],85:89]
			# bar<-paste(bar[which(!bar=="")],collapse=", ")
			# if(!bar==""){
				# foo<-paste(foo,", ",bar,sep="")
			# }
		# }	
	# short_author_vec[i]<-foo
# }
# talks_list[[1]]$matrix_author<-short_author_vec

# talk_mat<-matrix(nrow=6,ncol=7)
# head(talks)

# ### Read in draft from Courtney
# sched<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018.xlsx",sheetName="Sheet1",stringsAsFactors=F)

# ### Link up ABSTRACT IDs to Courtney's sheet ###
# colnames(sched)
# sched$"Entrant..Individual....ID"

# ### Split cleaned data frame by Entrant ID ###
# talks_entrant_ID<-split(talks,talks$"Entrant..Individual....ID")
# nsubs<-sapply(talks_entrant_ID,nrow)

# sched$"Entrant..Individual....ID"[!is.na(sched$"Entrant..Individual....ID") ][!sched$"Entrant..Individual....ID"[!is.na(sched$"Entrant..Individual....ID") ]%in% names(nsubs)]

# sched$Abstract_ID<-rep(NA,nrow(sched))

# for(i in 1:nrow(sched)){
	# if(is.na(sched[i,]$"Entrant..Individual....ID")){next}else{		
		# foo<-nsubs[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]
		# if(length(foo)==0){
			# sched$Abstract_ID[i]<-"HALP"
		# }else{
			# if(nsubs[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]==1){
				# sched$Abstract_ID[i]<-rownames(talks_entrant_ID[[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]])
			# }else{
				# sched$Abstract_ID[i]<-"HALP"			
				# }
			# }
		# }
# }

# sched$Abstract_ID

# write.xlsx(sched,"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_4manedits.xlsx")

# ### 

# sched$dayblock<-sapply(strsplit(sched$Day.Room.TimeCode,"[[:digit:]]"),function(x) x[length(x)])

# sched_list<-split(sched,sched$Day)
# names(sched_list)

# sched_list[[4]]<-split(sched_list[[4]],sched_list[[4]]$dayblock)

# test<-sched_list[[4]]["A"][[1]]

# test<-split(test,test$Room)

# lapply(test, function(x) x$Time)

# cat(paste(c("Room",names(test)), collapse= " & ")," \\\\",sep="")

# ### Get abstract IDs for formatted talks ###
# for(i in 1:length(test)){
	# foo<-talks_man[talks_man$"Entrant..Individual....ID" %in% test[[i]]$"Entrant..Individual....ID",]
	# foo<-foo[foo$submissionType == "15-min talk",]
	
	# if(length(table(foo$Invited.Talk))>1){
		# foo<-foo[-grep(names(which.min(table(foo$Invited.Talk))),foo$Invited.Talk),]
	# }
	
	# test[[i]]$Title<-foo[order(test[[i]]$"Entrant..Individual....ID"),]$Title
	
	# foo<-talks_list[[1]][talks_list[[1]]$"Entrant..Individual....ID" %in% test[[i]]$"Entrant..Individual....ID",]
	# test[[i]]$matrix_author<-foo$matrix_author
		
# }

# ### Read in manually matched abstract IDs ###
# sched_man<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_withmanedits.xlsx",sheetName="Sheet1",stringsAsFactors=F)

# colnames(talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,])

# sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Title<-talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,]$Title

# sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract<-talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,]$Abstract

# sched_man$matrix_author<-talks_list[[1]][sched_man$Abstract_ID,]$matrix_author

# sched_man$Title.1<-NULL
# sched_man$Abstract.1<-NULL

# sched_man[,grep("Topical[.]Area",colnames(sched_man))]<-NULL

# sched_man<-sched_man[order(sched$"Entrant..Individual....ID",sched$"GS.Title"),]

# save(sched_man,file="sched_man.Rdata")
# load("sched_man.Rdata")
# sched_man[,1]<-NULL
# write.xlsx2(sched_man,"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_forGoogle.xlsx",row.names=F)


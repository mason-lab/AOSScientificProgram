### Clean up excel file from MemberSuite for LaTeX and upload to Sched.doc ###
require(xlsx)
require(XML)
require(tools)
require(Hmisc)

### Read in xlsx ###
talks<-read.xlsx2("~/AOSScientificProgram/2019_Anchorage/AOS 2019 Abstracts_Master_28 May 2019.xlsx",stringsAsFactors=F,sheetName="Search Results")
talks$LightningTitle<-rep("",nrow(talks))

### Get lightning titles ###
lighttalks<-read.xlsx2("~/AOSScientificProgram/2019_Anchorage/AOS 2019 Abstracts_Master_21 May 2019.xlsx",stringsAsFactors=F,sheetName="Search Results")

lightID<-lighttalks[lighttalks$LightningTitle!="",]$Entry.ID
lighttite<-lighttalks[lighttalks$LightningTitle!="",]$LightningTitle

for(i in 1:length(lightID)){
	talks[which(talks$Entry.ID == lightID[i]),]$LightningTitle<-lighttite[i]
}

### Remove weird characters from Abstracts ###
talks$Abstract<-gsub("","---",talks$Abstract)
talks$Abstract <-gsub("\\n"," ",talks$Abstract)
talks$Abstract <-gsub("&","\\\\&",talks$Abstract)
talks$Abstract <-gsub("’","'",talks$Abstract)
talks$Abstract <-gsub("ʻ","'",talks$Abstract)
talks$Abstract <-gsub("“","\"",talks$Abstract)
talks$Abstract <-gsub("”","\"",talks$Abstract)
talks$Abstract <-gsub("‘","'",talks$Abstract)
talks$Abstract <-gsub("%","\\\\%",talks$Abstract)
talks$Abstract <-gsub("\\$","\\\\$",talks$Abstract)
talks$Abstract <-gsub("<","$<$",talks$Abstract)
talks$Abstract <-gsub(">","$>$",talks$Abstract)
talks$Abstract <-gsub("_","\\\\_",talks$Abstract)
talks$Abstract <-gsub("\\^","\\textasciicircum ",talks$Abstract)

### Remove weird characters from Title ###
talks$Title<-gsub("\\n"," ",talks$Title)
talks$Title<-gsub("&","\\\\&",talks$Title)
talks$Title<-gsub("’","'",talks$Title)
talks$Title<-gsub("“","\"",talks$Title)
talks$Title<-gsub("”","\"",talks$Title)
talks$Title<-gsub("ʻ","'",talks$Title)
talks$Title<-gsub("‘","'",talks$Title)
talks$Title<-gsub("Ō","\\={O}",talks$Title)
talks$Title<-gsub("–","--",talks$Title)
talks$Title<-gsub("–","--",talks$Title)
talks$Title<-gsub("°","\\textdegree ",talks$Title)

### Fix Session Titles ###
talks$Session.Title<-gsub("Sym","Symposium",talks$Session.Title)
talks$Session.Title<-gsub("’","'",talks$Session.Title)
talks$Session.Title<-gsub("Popn","Population",talks$Session.Title)

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
	talks[,firstlastcols[i]]<-gsub("ñ","\\\\~{n}", talks[,firstlastcols[i]])
	talks[,firstlastcols[i]] <-sapply(talks[,firstlastcols[i]], function(x) gsub("^\\s+|\\s+$", "", x))
}

### Create Cleaned Author Strings [both FULL and SHORT versions] ###
namecols<-grep("First.Name",colnames(talks))
name_vec_lists<-lapply(seq(namecols[2], namecols[length(namecols)],6),function(x) (x:(x+2)))

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

### Encapsulate titles in \capitalisewords for LaTeX ###
talks$Title<-paste0("\\capitalisewords{",talks$Title,"}")

### Create tex files for talk matrix ###
posters_ids<-talks[talks$FORMAT=="Poster",]
talks<-talks[!(talks$FORMAT=="Poster"),]

### Get Room Info for Talks ###
### RETURN TO THIS IF WE WANT TO CHANGE ROOM ORDER AND WE WANT to reordeR COLUMNS ACCORDINGLY ###
talks$Room.Name<-factor(talks$Room.Name,levels=levels(factor(talks$Room.Name))[c(1,2,3,4,7,8,9,10,11,12,13,5,6)])

### Set up Time Slots
time_slots<-list(TimeSetA=c("10:30","10:45","11:00","11:15","11:30","11:45"),TimeSetB=c("14:00","14:15","14:30","14:45","15:00","15:15"),TimeSetC=c("16:00","16:15","16:30","16:45","17:00","17:15"))

talks$TimeSession<-rep(NA,nrow(talks))
talks$TimeSession[talks$Time %in% gsub(":","",time_slots[[1]])]<-"A"
talks$TimeSession[talks$Time %in% gsub(":","",time_slots[[2]])]<-"B"
talks$TimeSession[talks$Time %in% gsub(":","",time_slots[[3]])]<-"C"


day_list<-split(talks,talks$Day)
day_time_list<-lapply(day_list,function(x) split(x,x$TimeSession))

### Create tex files for matrix, 2 per day (1 for rooms 1:6, another for rooms 7:12) ###
i<-2
j<-1
k<-1

for(i in 1:length(day_time_list)){
	for(j in 1:length(day_time_list[[i]])){
		day_time_list[[i]][[j]]$Room.Name<-factor(day_time_list[[i]][[j]]$Room.Name,levels=levels(day_time_list[[i]][[j]]$Room.Name)[unique(sort(as.numeric(day_time_list[[i]][[j]]$Room.Name)))])
		
		day_time_list[[i]][[j]]$Room.Set<-rep(NA,nrow(day_time_list[[i]][[j]]))

		day_time_list[[i]][[j]]$Room.Set[day_time_list[[i]][[j]]$Room.Name %in% levels(day_time_list[[i]][[j]]$Room.Name)[1:6]]<-"RoomSet1"
		day_time_list[[i]][[j]]$Room.Set[day_time_list[[i]][[j]]$Room.Name %in% levels(day_time_list[[i]][[j]]$Room.Name)[7:length(levels(day_time_list[[i]][[j]]$Room.Name))]]<-"RoomSet2"

		day_time_room_list<-split(day_time_list[[i]][[j]],day_time_list[[i]][[j]]$Room.Set)	
		
		for(k in 1:length(day_time_room_list)){
			day_time_room_list[[k]]$Room.Name<-factor(day_time_room_list[[k]]$Room.Name,levels=levels(day_time_room_list[[k]]$Room.Name)[unique(sort(as.numeric(day_time_room_list[[k]]$Room.Name)))])				

			### Ascertain which columns are symposia and color columns accordingly ###
			symp_rooms<-unique(as.numeric(day_time_room_list[[k]]$Room.Name)[day_time_room_list[[k]]$FORMAT=="Symposium"])
			col_color_vec<-rep(1,length(levels(day_time_room_list[[k]]$Room.Name)))
			col_color_vec[symp_rooms]<-2
			
			sink(file=paste("~/AOSScientificProgram/2019_Anchorage/BlockSchedule/",names(day_list)[i],"-TimeSlot",names(day_time_list[[i]])[j],"-",names(day_time_room_list)[k],".tex",sep=""))
			
			cat("\\begin{tabular}{|x{0.8cm}")
			for(m in 1:length(col_color_vec)){
				cat("|")
				cat(c("x","a")[col_color_vec[m]])
				cat("{2.65cm}")
			}
			cat("|@{}m{0pt}@{}}\\hline\n")
			
			cat("Room",levels(day_time_room_list[[k]]$Room.Name),sep=" & ")
			cat("&\\\\\n")
			cat("\\hline\n")
			
			### Deal with EP Mini Symposium Later ### 24 May 2019
			
			### Format LIGHTNING talks if present ###
			if(any(day_time_room_list[[k]]$Session.Title == "Lightning Talks")){
				light_time_slots<-unlist(lapply(as.numeric(unique(day_time_room_list[[k]]$Time)),function(x) seq(x,x+10,5)))
				light_tf<-talks$Day==day_time_room_list[[k]]$Day[1] & talks$TimeSession == day_time_room_list[[k]]$TimeSession[1] & talks$Time %in% light_time_slots & as.character(talks$Room.Name) %in% as.character(day_time_room_list[[k]]$Room.Name)
				light_tf[is.na(light_tf)]<-T
				
				talks[light_tf,]$Session.Title
				
				day_time_room_list[[k]]<-talks[light_tf,]
											
				light_titles<-paste0("\\scriptsize ", day_time_room_list[[k]][day_time_room_list[[k]] $Session.Title == "Lightning Talks",]$LightningTitle,"\\par \\tiny ",day_time_room_list[[k]][day_time_room_list[[k]] $Session.Title == "Lightning Talks",]$"Author 1 Short Name"," et al. ")
				
				light_titles_full<-sapply(lapply(seq(1,13,3),function(x) x:(x+2)),function(x) paste(light_titles[x],collapse="\\par - - - - - - - - - - - - - - - - - \\par \\vspace{2pt} "))
				day_time_room_list[[k]][day_time_room_list[[k]]$Session.Title== "Lightning Talks",]$Title[seq(1,13,3)]<-light_titles_full
				}

			day_time_room_list[[k]]<-day_time_room_list[[k]][!is.na(day_time_room_list[[k]]$TimeSession),]
			
			times<-unique(day_time_room_list[[k]]$Time)
			
			for(m in 1:length(times)){
				this_time<-day_time_room_list[[k]][day_time_room_list[[k]]$Time == times[m],]				
					
				this_time<-this_time[order(as.numeric(this_time$Room.Name)),]
				rownames(this_time)<-this_time$Room.Name
				this_time<-this_time[levels(day_time_room_list[[k]]$Room.Name)[sort(as.numeric(unique(day_time_room_list[[k]]$Room.Name)))],]
				
				if(length(grep("NA",rownames(this_time)))>0){
					this_time[grep("NA",rownames(this_time)),]<-rep("",ncol(this_time))
					rownames(this_time)<-levels(day_time_room_list[[k]]$Room.Name)[sort(as.numeric(unique(day_time_room_list[[k]]$Room.Name)))]
				}
				### Make session header for this page ###	
				if(m==1){
						cat("\\rule{0pt}{1em} ")			
						cat("\\textbf{Session} &")
						symptite<-gsub("/","\\textbackslash ",gsub("&","\\\\&", this_time$Session.Title),fixed=T)
						symptite<-paste("\\footnotesize \\textbf{\\capitalisewords{",symptite,"}}",sep="")
						cat(symptite, sep=" & ")
						cat("&\\\\[25ex]\n")
						#cat("\\rule{0pt}{1em} ")			
						cat("\\hline\n")
				}
				
				### Write out time in first column ### 
				cat("\\makecell{",times[m],"}&",sep="")
							
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
						cat(this_time$Title[n]," \\par \\vspace{8pt} ", "\\textit{", this_time$FullShortAuthor[n],"}",sep="")
						}
					if(n<nrow(this_time)){cat(" & ")}else{next}
				}
			cat("&\\\\[25ex]\n\\hline\n")
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
sink()
i
j
k

### Create Poster Output ###
posters<-posters_ids
posters$Day ## Created earlier in this same file in which we split talks from posters ###

posters$Poster.Number<-paste(sapply(strsplit(posters$Day,""),function(x) x[1]),".",sprintf("%03s",posters$Topical.Poster.Order),sep="")
posters<-posters[order(posters$Poster.Number),]

### Split posters into two different days ###
poster_sessions<-split(posters,posters$Day)

for(i in 1:length(poster_sessions)){
	
	sink(paste("~/AOSScientificProgram/2019_Anchorage/PosterList/",c("wednesday","thursday")[i],".tex",sep=""))
	
	poster_sessions[[i]]$PaperPoster[is.na(poster_sessions[[i]]$PaperPoster)]<-"0"
	poster_sessions[[i]]$Student.Prez.Award.Competitors.1[is.na(poster_sessions[[i]]$Student.Prez.Award.Competitors.1)]<-"0"

	for(j in 1:nrow(poster_sessions[[i]])){
		if(as.character(poster_sessions[[i]]$PaperPoster)[j]=="1"){
			cat("\\posterentry{",poster_sessions[[i]]$Poster.Number[j],"}{",sep="")
	
			if(poster_sessions[[i]]$Student.Prez.Award.Competitors.1[j]==1){
				poster_sessions[[i]]$FullShortAuthor[j]<-paste("*", poster_sessions[[i]]$FullShortAuthor[j],sep="")
			}
			
			cat(poster_sessions[[i]]$FullShortAuthor[j],"}{",poster_sessions[[i]]$Title[j],"}\n",sep="")
		}else{
			next}
		}	
	sink()
}

########################
### ABSTRACT BOOKLET ###
########################
posters_ab<-posters

### Sort by last name of presenting author ###
posters_which_auth<-as.numeric(sapply(strsplit(posters_ab $Presenting.Author,""),function(x) x[1]))
posters_column_vec<-seq(60,126,6)[posters_which_auth]

poster_present_vec<-vector()
for(i in 1:length(posters_which_auth)){
	poster_present_vec[i]<-posters_ab[i, posters_column_vec[i]]
	auth_foo<-strsplit(posters_ab $FullLongAuthor[i],", ")[[1]]
	auth_foo[posters_which_auth[i]] <-paste0("\\underline{",auth_foo[posters_which_auth[i]],"}")
	posters_ab $FullLongAuthor[i]<-paste(auth_foo,collapse=", ")
}
posters_ab <-posters_ab[order(poster_present_vec),]

sink("~/AOSScientificProgram/2019_Anchorage/AbstractBook/posters.tex")
for(i in 1:nrow(posters_ab)){
	cat("\\normaltalk")
	cat("{", posters_ab $Title[i],"}",sep="")
	cat("{", posters_ab $FullLongAuthor[i],"}",sep="")
	cat("{", posters_ab $Abstract[i],"}",sep="")
	cat("\n\n")
}
sink()

### Lightning Talks ###
l_talks<-talks[talks$Session.Title=="Lightning Talks",]
l_talks<-l_talks[!l_talks$Presenting.Author=="",]

### Sort by last name of presenting author ###
which_auth<-as.numeric(sapply(strsplit(l_talks$Presenting.Author,""),function(x) x[1]))
present_vec<-seq(60,126,6)[which_auth]

ln_present_vec<-vector()
for(i in 1:length(present_vec)){
	ln_present_vec[i]<-l_talks[i,present_vec[i]]
	auth_foo<-strsplit(l_talks$FullLongAuthor[i],", ")[[1]]
	auth_foo[which_auth[i]] <-paste0("\\underline{",auth_foo[which_auth[i]],"}")
	l_talks$FullLongAuthor[i]<-paste(auth_foo,collapse=", ")
}
l_talks<-l_talks[order(ln_present_vec),]

sink("~/AOSScientificProgram/2019_Anchorage/AbstractBook/lightningtalks.tex")
for(i in 1:nrow(l_talks)){
	cat("\\normaltalk")
	cat("{", l_talks $Title[i],"}",sep="")
	cat("{", l_talks $FullLongAuthor[i],"}",sep="")
	cat("{", l_talks $Abstract[i],"}",sep="")
	cat("\n\n")
}

sink()

### "NORMAL" Talks ###
n_talks<-talks[talks$Entry.ID!="",]

### Sort by last name of presenting author ###
n_talks_which_auth<-as.numeric(sapply(strsplit(n_talks$Presenting.Author,""),function(x) x[1]))
n_talks_present_vec<-seq(60,126,6)[n_talks_which_auth]

n_talks_present_author<-vector()
for(i in 1:length(n_talks_present_vec)){
	n_talks_present_author[i]<-n_talks[i, n_talks_present_vec[i]]
	auth_foo<-strsplit(n_talks $FullLongAuthor[i],", ")[[1]]

	auth_foo[n_talks_which_auth[i]] <-paste0("\\underline{",auth_foo[n_talks_which_auth[i]],"}")
	n_talks$FullLongAuthor[i]<-paste(auth_foo,collapse=", ")
}
n_talks <-n_talks[order(n_talks_present_author),]

sink("~/AOSScientificProgram/2019_Anchorage/AbstractBook/normaltalks.tex")
for(i in 1:nrow(n_talks)){
	cat("\\normaltalk")
	cat("{", n_talks $Title[i],"}",sep="")
	cat("{", n_talks $FullLongAuthor[i],"}",sep="")
	cat("{", n_talks $Abstract[i],"}",sep="")
	cat("\n\n")
}
sink()

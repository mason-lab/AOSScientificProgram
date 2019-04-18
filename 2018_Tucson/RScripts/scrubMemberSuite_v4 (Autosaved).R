### Clean up excel file from MemberSuite for LaTeX and upload to Sched.doc ###
require(xlsx)
require(XML)
require(tools)
require(Hmisc)

### Convenience functions ###
html2txt <- function(str) {
      xpathApply(htmlParse(str, asText=TRUE,trim=TRUE),
                 "//body//text()", 
                 xmlValue)[[1]] 
}

talks<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/Submissions_2-28-18.xlsx",sheetName="Search Results",stringsAsFactors=F,row.names=1)
head(talks)

head(talks)

### See format of data frame here ###
colnames(talks)

### Remove HTML code and other weird characters  from Abstracts ###
talks$Abstract<-gsub("<.*?>","",talks$Abstract)
talks$Abstract<-gsub("[{].*?[}]","",talks$Abstract)
talks$Abstract<-gsub("\n"," ",talks$Abstract)
talks$Abstract<-gsub("p.p1","",talks$Abstract)
talks$Abstract<-sapply(talks$Abstract, html2txt)
talks$Abstract<-gsub(".*?table.MsoNormalTable      ","",talks$Abstract)
talks$Abstract<-gsub(".*?p.ctl    ","",talks$Abstract)
talks$Abstract<-gsub(".*?   X-NONE","",talks$Abstract)
talks$Abstract<-gsub(".*?span.Apple-tab-span    ","",talks$Abstract)
talks$Abstract<-gsub(".*?p   ","",talks$Abstract)
talks$Abstract<-sapply(talks$Abstract, function(x) gsub("^\\s+|\\s+$", "", x))


### Remove HTML code and weird characters from Title ###
talks$Title<-gsub("<.*?>","",talks$Title)
talks$Title<-gsub("[{].*?[}]","",talks$Title)
talks$Title<-gsub("\n"," ",talks$Title)
talks$Title<-gsub("p.p1","",talks$Title)
talks$Title<-sapply(talks$Title, html2txt)
talks$Title<-gsub(".*?table.MsoNormalTable","",talks$Title)
talks$Title<-gsub(".*?p.ctl    ","",talks$Title)
talks$Title<-gsub(".*?   X-NONE","",talks$Title)
talks$Title<-gsub(".*?span.Apple-tab-span    ","",talks$Title)
talks$Title<-gsub(".*?p   ","",talks$Title)
talks$Title<-sapply(talks$Title, function(x) gsub("^\\s+|\\s+$", "", x))

for(i in 1:length(talks$Title)){
	if(grepl("\\.",strsplit(talks$Title[i],"")[[1]][length(strsplit(talks$Title[i],"")[[1]])])){ #Removes periods at end
		talks$Title[i]<-paste(strsplit(talks$Title[i],"")[[1]][1:(length(strsplit(talks$Title[i],"")[[1]])-1)],collapse="")
	}
	if(grepl(" ",strsplit(talks$Title[i],"")[[1]][length(strsplit(talks$Title[i],"")[[1]])])){
		talks$Title[i]<-paste(strsplit(talks$Title[i],"")[[1]][1:(length(strsplit(talks$Title[i],"")[[1]])-1)],collapse="")		
	}
}	

talks$Title[talks$Title==toupper(talks$Title)]<-toTitleCase(gsub("([[:alpha:]])([[:alpha:]]+)", "\\U\\1\\L\\2", talks$Title[talks$Title==toupper(talks$Title)], perl=TRUE))

# write.xlsx(talks[noabstract,],file="/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_noabstract_28Feb2018.xlsx")
# write.xlsx(talks[notitle,],file="/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_notitle_28Feb2018.xlsx")
# write.xlsx(talks[notitle_noabstract,],file="/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_notitle_noabstract_28Feb2018.xlsx")

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

### Create Cleaned Author Strings [both FULL and SHORT versions] ###
colnames(talks)
colnames(talks)[23]
talks[23]=="Yes"

name_vec_lists<-list()
name_vec_lists[[1]]<-c(4,5,6)
name_vec_lists[[2]]<-c(24,25,26)
name_vec_lists[[3]]<-c(30,31,32)
name_vec_lists[[4]]<-c(36,37,38)
name_vec_lists[[5]]<-c(42,43,44)
name_vec_lists[[6]]<-c(48,49,50)
name_vec_lists[[7]]<-c(54,55,56)

### Create long version names ###
for(i in 1:length(name_vec_lists)){
	talks[,(ncol(talks)+1)]<-rep(NA,nrow(talks))
	for(j in 1:nrow(talks)){
		talks[j,ncol(talks)]<-gsub("\\s+"," ",paste(talks[j,name_vec_lists[[i]]],collapse=" "))
	}
}

colnames(talks)[(ncol(talks)-6):ncol(talks)]<-c("Presenting Author Full Name",paste("Author",1:6,"Full Name"))

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

head(talks)

### Combined Short Author Fields ###
#talks[,(ncol(talks)+1)]<-rep("",nrow(talks))
# sn_vec<-(ncol(talks)-6):ncol(talks)

# talks[i, sn_vec]
# foo<-rep("",nrow(talks))
# for(i in 1:nrow(talks)){
	# if(talks[i,23]=="Yes"){
		# bar<-gsub(" ,","",paste(talks[i,sn_vec[-2]],collapse=", "))
		# bar2<-strsplit(bar,"")[[1]]
		# if(paste(bar2[(length(bar2)-1):length(bar2)],collapse="")==", "){
			# foo[i]<-paste(strsplit(bar,"")[[1]][1:(length(strsplit(bar,"")[[1]])-2)],collapse="")
			# }else{
			# foo[i]<-bar
				# }
			# }else{
		# foo[i]<-gsub(" ,","",paste(talks[i,sn_vec[-1]],collapse=", "))
		# foo[i]<-paste(strsplit(foo[i],"")[[1]][1:(length(strsplit(foo[i],"")[[1]])-2)],collapse="")
	# if(foo[i]==","){
		# bar<-gsub(" ,","",paste(talks[i,sn_vec],collapse=", "))
		# foo[i]<-paste(strsplit(bar,"")[[1]][1:(length(strsplit(bar,"")[[1]])-2)],collapse="")
		# }
	# }
# }
# foo

# talks$AuthorsCondensed<-foo

### Format 15-min / lightning / poster fields ###
rank<-grep("Rank",colnames(talks))

for(i in 1:length(rank)){
	talks[,rank[i]]<-gsub("Do not want this format","0",talks[,rank[i]])
	talks[,rank[i]]<-gsub("No preference","0",talks[,rank[i]])
	talks[,rank[i]]<-gsub("1st Choice","1",talks[,rank[i]])
	talks[,rank[i]]<-gsub("2nd Choice","2",talks[,rank[i]])
	talks[,rank[i]]<-gsub("3rd Choice","3",talks[,rank[i]])
}

submissionType<-apply(talks[,rank],1,function(x) c("15-min talk","Lightning Talk","Poster")[which(x=="1")])
submissionType[which(sapply(submissionType, function(x) identical(x,character(0))))]<-"15-min talk"

talks$submissionType<-unlist(submissionType)

### Create table with relevant outputs to manually clean ###
colnames(talks)

talks_cleaned<-talks[c(90,75,76,7,77,27,78,33,79,39,80,45,81,51,82,57,23,71,72,61,62,73)]

write.xlsx(talks_cleaned, "/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_Submissions_Cleaned_01Mar2018_NAM.xlsx")

### Read in Cleaned XLSX ###
talks_man<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/CleanedMSOutput/Tucson2018_Submissions_Cleaned_01Mar2018_NAM.xlsx",sheetName="Sheet1",row.names=1,stringsAsFactors=F)
talks_man[is.na(talks_man)]<-""
talks_man[talks_man==" "]<-""

aff_talks_man<-grep("Affiliation",colnames(talks_man))

for(i in 1:length(aff_talks_man)){
	talks_man[,aff_talks_man[i]]<-gsub("&","\\\\&",talks_man[,aff_talks_man[i]])
}

talks_man$Abstract<-gsub("’","'",talks_man$Abstract)
talks_man$Abstract<-gsub("~","{\\raise.17ex\\hbox{$\\scriptstyle\\mathtt{\\sim}$}}",talks_man$Abstract)
talks_man$Title<-gsub("’","'",talks_man$Title)
talks_man$Title<-gsub("&","\\\\&", talks_man$Title)

talks_man$Abstract<-latexTranslate(talks_man$Abstract)

### Split into different talk types ###
talks_list<-split(talks_man,talks_man$submissionType)

### NORMAL TALKS ###
### Order according to last name of presenting author ###
talks_list[[1]]<-talks_list[[1]][order(talks[rownames(talks_list[[1]]),6]),]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/oralpresentations.tex")
for(i in 1:nrow(talks_list[[1]])){
	cat("\\normaltalk")
	if(talks_list[[1]][i,17]=="Yes"){
		cat("{",talks_list[[1]][i,3],", ",talks_list[[1]][i,4],"}",sep="")
	}else{
		cat("{",talks_list[[1]][i,5],", ",talks_list[[1]][i,6],"}",sep="")
	}
	cat("{",talks_list[[1]][i,7],", ",talks_list[[1]][i,8],"}",sep="")
	cat("{",talks_list[[1]][i,9],", ",talks_list[[1]][i,10],"}",sep="")
	cat("{",talks_list[[1]][i,11],", ",talks_list[[1]][i,12],"}",sep="")
	cat("{",talks_list[[1]][i,13],", ",talks_list[[1]][i,14],"}",sep="")
	cat("{",talks_list[[1]][i,15],", ",talks_list[[1]][i,16],"}",sep="")

	cat("{",talks_list[[1]][i,20],"}",sep="")			
	cat("{",talks_list[[1]][i,21]," [",rownames(talks_list[[1]])[i],"]}",sep="")	
			
	cat("\n\n")
}

sink()

### LIGHTNING TALKS ###
### Order according to last name of presenting author ###
talks_list[[2]]<-talks_list[[2]][order(talks[rownames(talks_list[[2]]),6]),]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/lightningtalks.tex")
for(i in 1:nrow(talks_list[[2]])){
	cat("\\normaltalk")
	if(talks_list[[2]][i,17]=="Yes"){
		cat("{",talks_list[[2]][i,3],", ",talks_list[[2]][i,4],"}",sep="")
	}else{
		cat("{",talks_list[[2]][i,5],", ",talks_list[[2]][i,6],"}",sep="")
	}
	cat("{",talks_list[[2]][i,7],", ",talks_list[[2]][i,8],"}",sep="")
	cat("{",talks_list[[2]][i,9],", ",talks_list[[2]][i,10],"}",sep="")
	cat("{",talks_list[[2]][i,11],", ",talks_list[[2]][i,12],"}",sep="")
	cat("{",talks_list[[2]][i,13],", ",talks_list[[2]][i,14],"}",sep="")
	cat("{",talks_list[[2]][i,15],", ",talks_list[[2]][i,16],"}",sep="")

	cat("{",talks_list[[2]][i,20],"}",sep="")			
	cat("{",talks_list[[2]][i,21]," [",rownames(talks_list[[2]])[i],"]}",sep="")	
			
	cat("\n\n")
}

sink()

### POSTERS ###
### Order according to last name of presenting author ###
talks_list[[3]]<-talks_list[[3]][order(talks[rownames(talks_list[[3]]),6]),]

sink("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/AbstractBooklet/posters.tex")
for(i in 1:nrow(talks_list[[3]])){
	cat("\\normaltalk")
	if(talks_list[[3]][i,17]=="Yes"){
		cat("{",talks_list[[3]][i,3],", ",talks_list[[3]][i,4],"}",sep="")
	}else{
		cat("{",talks_list[[3]][i,5],", ",talks_list[[3]][i,6],"}",sep="")
	}
	cat("{",talks_list[[3]][i,7],", ",talks_list[[3]][i,8],"}",sep="")
	cat("{",talks_list[[3]][i,9],", ",talks_list[[3]][i,10],"}",sep="")
	cat("{",talks_list[[3]][i,11],", ",talks_list[[3]][i,12],"}",sep="")
	cat("{",talks_list[[3]][i,13],", ",talks_list[[3]][i,14],"}",sep="")
	cat("{",talks_list[[3]][i,15],", ",talks_list[[3]][i,16],"}",sep="")

	cat("{",talks_list[[3]][i,20],"}",sep="")			
	cat("{",talks_list[[3]][i,21]," [",rownames(talks_list[[3]])[i],"]}",sep="")	
			
	cat("\n\n")
}

sink()

### Generate Output for Schedule Matrix Template ###
short_author_vec<-vector()
colnames(talks)

for(i in 1:nrow(talks_list[[1]])){
	colnames(talks_list[[1]])
	if(talks_list[[1]][i,17]=="Yes"){
		foo<-talks[rownames(talks_list[[1]])[i],83]
		bar<-talks[rownames(talks_list[[1]])[i],85:89]
		bar<-paste(bar[which(!bar=="")],collapse=", ")
		if(!bar==""){
			foo<-paste(foo,", ",bar,sep="")
			}
		}else{
			foo<-talks[rownames(talks_list[[1]])[i],84]
			bar<-talks[rownames(talks_list[[1]])[i],85:89]
			bar<-paste(bar[which(!bar=="")],collapse=", ")
			if(!bar==""){
				foo<-paste(foo,", ",bar,sep="")
			}
		}	
	short_author_vec[i]<-foo
}
talks_list[[1]]$matrix_author<-short_author_vec

talk_mat<-matrix(nrow=6,ncol=7)
head(talks)

### Read in draft from Courtney
sched<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018.xlsx",sheetName="Sheet1",stringsAsFactors=F)

### Link up ABSTRACT IDs to Courtney's sheet ###
colnames(sched)
sched$"Entrant..Individual....ID"

### Split cleaned data frame by Entrant ID ###
talks_entrant_ID<-split(talks,talks$"Entrant..Individual....ID")
nsubs<-sapply(talks_entrant_ID,nrow)

sched$"Entrant..Individual....ID"[!is.na(sched$"Entrant..Individual....ID") ][!sched$"Entrant..Individual....ID"[!is.na(sched$"Entrant..Individual....ID") ]%in% names(nsubs)]

sched$Abstract_ID<-rep(NA,nrow(sched))

for(i in 1:nrow(sched)){
	if(is.na(sched[i,]$"Entrant..Individual....ID")){next}else{		
		foo<-nsubs[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]
		if(length(foo)==0){
			sched$Abstract_ID[i]<-"HALP"
		}else{
			if(nsubs[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]==1){
				sched$Abstract_ID[i]<-rownames(talks_entrant_ID[[which(names(nsubs)==sched[i,]$"Entrant..Individual....ID")]])
			}else{
				sched$Abstract_ID[i]<-"HALP"			
				}
			}
		}
}

sched$Abstract_ID

write.xlsx(sched,"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_4manedits.xlsx")

### 

sched$dayblock<-sapply(strsplit(sched$Day.Room.TimeCode,"[[:digit:]]"),function(x) x[length(x)])

sched_list<-split(sched,sched$Day)
names(sched_list)

sched_list[[4]]<-split(sched_list[[4]],sched_list[[4]]$dayblock)

test<-sched_list[[4]]["A"][[1]]

test<-split(test,test$Room)

lapply(test, function(x) x$Time)

cat(paste(c("Room",names(test)), collapse= " & ")," \\\\",sep="")

### Get abstract IDs for formatted talks ###
for(i in 1:length(test)){
	foo<-talks_man[talks_man$"Entrant..Individual....ID" %in% test[[i]]$"Entrant..Individual....ID",]
	foo<-foo[foo$submissionType == "15-min talk",]
	
	if(length(table(foo$Invited.Talk))>1){
		foo<-foo[-grep(names(which.min(table(foo$Invited.Talk))),foo$Invited.Talk),]
	}
	
	test[[i]]$Title<-foo[order(test[[i]]$"Entrant..Individual....ID"),]$Title
	
	foo<-talks_list[[1]][talks_list[[1]]$"Entrant..Individual....ID" %in% test[[i]]$"Entrant..Individual....ID",]
	test[[i]]$matrix_author<-foo$matrix_author
		
}

### Read in manually matched abstract IDs ###
sched_man<-read.xlsx2("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_withmanedits.xlsx",sheetName="Sheet1",stringsAsFactors=F)

colnames(talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,])

sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Title<-talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,]$Title

sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract<-talks_man[sched_man[!(is.na(sched_man$Abstract_ID) | sched_man$Abstract_ID %in% "?????"),]$Abstract_ID,]$Abstract

sched_man$matrix_author<-talks_list[[1]][sched_man$Abstract_ID,]$matrix_author

sched_man$Title.1<-NULL
sched_man$Abstract.1<-NULL

sched_man[,grep("Topical[.]Area",colnames(sched_man))]<-NULL

sched_man<-sched_man[order(sched$"Entrant..Individual....ID",sched$"GS.Title"),]

save(sched_man,file="sched_man.Rdata")
load("sched_man.Rdata")
sched_man[,1]<-NULL
write.xlsx2(sched_man,"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/PresentationSchedule_05March2018_forGoogle.xlsx",row.names=F)


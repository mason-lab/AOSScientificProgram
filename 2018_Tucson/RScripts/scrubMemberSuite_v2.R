### Clean up excel file from MemberSuite for LaTeX and upload to Sched.doc ###
require(xlsx)
require(XML)
require(tools)

### Convenience functions ###
html2txt <- function(str) {
      xpathApply(htmlParse(str, asText=TRUE,trim=TRUE),
                 "//body//text()", 
                 xmlValue)[[1]] 
}

talks<-read.xlsx("~/Desktop/Service/AOS_ProgramBooklet/Tucson2018/Samples/Submissions_2-28-18.xlsx",sheetName="Search Results",stringsAsFactors=F,row.names=1)
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
name_vec_lists[[4]]<-c(42,43,44)
name_vec_lists[[5]]<-c(48,49,50)
name_vec_lists[[6]]<-c(54,55,56)

for(i in 1:length(name_vec_lists)){
	talks[,(ncol(talks)+1)]<-rep(NA,nrow(talks))
	for(j in 1:nrow(talks)){
		talks[j,ncol(talks)]<-gsub("\\s+"," ",paste(talks[j,name_vec_lists[[i]]],collapse=" "))
	}
}

colnames(talks)[(ncol(talks)-5):ncol(talks)]<-c("Presenting Author Full Name",paste("Author",1:5,"Full Name"))

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

colnames(talks)[(ncol(talks)-5):ncol(talks)]<-c("Presenting Author Short Name",paste("Author",1:5,"Short Name"))
head(talks)

### Combined Short Author Fields ###
#talks[,(ncol(talks)+1)]<-rep("",nrow(talks))
sn_vec<-(ncol(talks)-5):ncol(talks)

talks[i, sn_vec]
foo<-rep("",nrow(talks))
for(i in 1:nrow(talks)){
	if(talks[i,23]=="Yes"){
		bar<-gsub(" ,","",paste(talks[i,sn_vec[-2]],collapse=", "))
		bar2<-strsplit(bar,"")[[1]]
		if(paste(bar2[(length(bar2)-1):length(bar2)],collapse="")==", "){
			foo[i]<-paste(strsplit(bar,"")[[1]][1:(length(strsplit(bar,"")[[1]])-2)],collapse="")
			}else{
			foo[i]<-bar
				}
			}else{
		foo[i]<-gsub(" ,","",paste(talks[i,sn_vec[-1]],collapse=", "))
		foo[i]<-paste(strsplit(foo[i],"")[[1]][1:(length(strsplit(foo[i],"")[[1]])-2)],collapse="")
	if(foo[i]==","){
		bar<-gsub(" ,","",paste(talks[i,sn_vec],collapse=", "))
		foo[i]<-paste(strsplit(bar,"")[[1]][1:(length(strsplit(bar,"")[[1]])-2)],collapse="")
		}
	}
}
talks$AuthorsCondensed<-foo

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

### Split into different submission types ###
subs<-split(talks,talks$submissionType)

### Create output for Abstract booklet ###
foo<-subs[[1]][,26]


as.factor(talks$submissionType)
### Write out cleaned excel file for da crew ###
colnames(talks)
talks_red<-talks[c(75,1,88,61,62,66,67,68,72)]
talks_red<-cbind(AbstractID=rownames(talks_red),talks_red)

write.xlsx(talks_red,"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_Submissions_Cleaned_28Feb2018_NAM.xlsx",row.names=F)

### Generate output for Abstract Booklet ###
### Find instances where same author submitted multiple abstracts ###
dup<-duplicated(talks$"Entrant (Individual) - ID") | duplicated(talks$"Entrant (Individual) - ID",fromLast=T)
talks_dup<-talks[dup,]

write.xlsx(talks[dup,],"/Users/NickMason/Desktop/Service/AOS_ProgramBooklet/Tucson2018/MissingInfo/Tucson2018_duplicated_v2.xlsx")

### Create cleaned data frame(s) ###

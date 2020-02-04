#url <- "http://www.omegahat.net/R/bin/windows/contrib/3.5.1/RDCOMClient_0.93-0.zip"
#install.packages(url, repos=NULL, type="binary")

#set the working directory
setwd("C:/Users/gssaruba/Documents")

#loding packages
library("dbplyr",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("RMariaDB",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("dplyr",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("stringi",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("ggplot2",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("readxl",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("readr",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("ggthemes",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("ggExtra",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")
library("patchwork",lib.loc="C:/Users/gssaruba/Documents/R/R-3.6.1/library")

library(dbplyr)
library(RMariaDB)
library(dplyr)
library(stringi)
library(ggplot2)
library(readxl)
library(readr)
library(ggthemes)
library(ggExtra)
library(patchwork)



#Connecting to MySQL database
con <- DBI::dbConnect(RMariaDB::MariaDB(), 
                      dbname="etspro",
                      host = "GSSNT803.asia.ad.flextronics.com",
                      user = "business_excell",
                      password = "Chennai"
                      
)

#Preparing a new csv file based on Ibox approval conditions

current_data<-data.frame(dbGetQuery(con, "SELECT * FROM ideatab WHERE (idea_date >='2019-10-01')"))
current_data$project_status<-as.factor(current_data$project_status)
current_data$idea_status<-as.factor(current_data$idea_status)

#levels(current_data$idea_status)
#levels(current_data$project_status)<-c("Approved","Not Considered","Pending","NA")

r1<-current_data%>%filter(idea_status=='Approved' & project_status=='Approved')%>%mutate(new_status='Completed')
r2<-current_data%>%filter(idea_status=='Approved' & project_status=='Pending')%>%mutate(new_status='Inprogress')
r3<-current_data%>%filter(idea_status=='Approved' & project_status=='Not Considered')%>%mutate(new_status='Not Considered')
r4<-current_data%>%filter(idea_status=='Approved' & !project_status %in% c("Approved","Pending","Not Considered"))%>%mutate(new_status='Inprogress')

r5<-current_data%>%filter(idea_status=='Not Considered')%>%mutate(new_status='Not Considered')
r6<-current_data%>%filter(idea_status=='Pending'& !project_status %in% c("Approved","Pending","Not Considered"))%>%mutate(new_status='Pending')
new_data<-as.data.frame(rbind(r1,r2,r3,r4,r5,r6))

#Converting diff.names of PM vertical to single value
new_data$vertical<-as.factor(new_data$vertical)
new_data$vertical[new_data$vertical=='GBS Program Management']<-'GBS PM'
new_data$vertical[new_data$vertical=='GBS Program Mgmt']<-'GBS PM'


#Writing it as a new csv file
getwd()
write.csv(new_data,file="ibox_data.csv")

#Disconnecting it from MySQL
dbDisconnect(con)

#Reading workday profile csv from common share path
Emp_wkday<- read_excel("Z:/Quality/BUSINESS EXCELLENCE/GBS Chennai/Reports & Presentations/Idea Box/workday_emp_data/Employee_data_workday.xlsx")
Emp_wkday<-as.data.frame(Emp_wkday)

#Reading previously stored ibox file
ibox<- read_csv("ibox_data.csv",col_types = cols(X1 = col_skip()))
ibox<-as.data.frame(ibox)
ibox%>%group_by(format(as.Date(ibox$idea_date),"%Y"))%>%count()

#Preprocessing of Employee Workday file - Merge of Ibox with workday and conversion of levels of location
names(Emp_wkday)[names(Emp_wkday)== 'Employee ADID']<-'emp_adid'
names(Emp_wkday)[names(Emp_wkday)== 'Employee ID']<-'wd_id'
ibox_emp_m<-merge(x=ibox,y=Emp_wkday,by="wd_id")
names(ibox_emp_m)[names(ibox_emp_m)== 'Management Chain - Level 07']<-'Management_Chain_Level07'

ibox_emp_m$`Location Name`<-as.factor(ibox_emp_m$`Location Name`)
levels(ibox_emp_m$`Location Name`)<-c(levels(ibox_emp_m$`Location Name`),"GBS_Chennai","GBS_Pune")
levels(ibox_emp_m$`Location Name`)
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Chennai GBS B1']<-'GBS_Chennai'
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Chennai GBS B2']<-'GBS_Chennai'
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Chennai GBS B3']<-'GBS_Chennai'
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Chennai GBS B4']<-'GBS_Chennai'
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Pune SEZ B1']<-'GBS_Pune'
ibox_emp_m$`Location Name`[ibox_emp_m$`Location Name`=='IN Pune SEZ B2']<-'GBS_Pune'

ibox_emp_m$vertical<-as.factor(ibox_emp_m$vertical)
ibox_emp_m$idea_date<-as.Date(ibox_emp_m$idea_date)


ibox_emp_m$month<-months(ibox_emp_m$idea_date,abbreviate = TRUE)
ibox_emp_m$month<-factor(ibox_emp_m$month,levels=month.abb)
head(ibox_emp_m$idea_date)
class(ibox_emp_m$idea_date)

#Important : Filtering only for FY20 data
ibox_emp_m<-ibox_emp_m%>%filter(idea_date >'2019-03-31')



##-----------------------------Ibox Monthly trend---------------------------------------------------
library(zoo)
#Plot of Ibox Monthly Trend - GBS Chennai and Pune
g1<-ibox_emp_m%>%select(idea_date,month,ideano)%>%group_by(month=format(as.Date(ibox_emp_m$idea_date),"%Y-%m"))%>%summarise(count_ideas=n())
g1$month<-as.yearmon(g1$month)
g1$month<-as.Date(g1$month)


p1<-ggplot(data=g1,aes(x=g1$month,y=g1$count_ideas),fill=col)+geom_line(color="dodgerblue2")+geom_point(color="black",size=2)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+geom_text(label=g1$count_ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)+theme(axis.text.x = element_text(angle=0, hjust = 1))+ggtitle("Monthly Trend of Ideas Generated")+xlab("Months")+ylab("Numbers of Ideas Posted")+removeGrid()
p1


#Plot of Ibox Monthly Trend - GBS Chennai and Pune - GBS Engineering
g1<-ibox_emp_m%>%select(idea_date,month,ideano,`Location Name`,vertical)%>%group_by(month=format(as.Date(ibox_emp_m$idea_date),"%Y-%m"),`Location Name`,vertical)%>%filter(vertical=="GBS Engineering")%>%summarise(count_ideas=n())
g1$month<-as.yearmon(g1$month)
g1$month<-as.Date(g1$month)
p11<-ggplot(data=g1,aes(x=month,y=count_ideas),fill=col)+geom_line(color="dodgerblue2")+geom_point(color="black",size=2)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+geom_text(label=g1$count_ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)+theme(axis.text.x = element_text(angle=0, hjust = 1))+ggtitle("Monthly Trend of Ideas Generated - GBS Engineering")+xlab("Months")+ylab("Numbers of Ideas Posted")+removeGrid()+facet_wrap(~`Location Name`)
p11

#Plot of Ibox Monthly trend - GBS Chennai and Pune - GBS Supply Chain

g1<-ibox_emp_m%>%select(idea_date,month,ideano,`Location Name`,vertical)%>%group_by(month=format(as.Date(ibox_emp_m$idea_date),"%Y-%m"),`Location Name`,vertical)%>%filter(vertical=="GBS Supply Chain")%>%summarise(count_ideas=n())
g1$month<-as.yearmon(g1$month)
g1$month<-as.Date(g1$month)
p12<-ggplot(data=g1,aes(x=month,y=count_ideas),fill=col)+geom_line(color="dodgerblue2")+geom_point(color="black",size=2)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+geom_text(label=g1$count_ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)+theme(axis.text.x = element_text(angle=0, hjust = 1))+ggtitle("Monthly Trend of Ideas Generated - GBS Supply Chain")+xlab("Months")+ylab("Numbers of Ideas Posted")+removeGrid()+facet_wrap(~`Location Name`)
p12

#Plot of Ibox Monthly trend - GBS Chennai and Pune - GBS Finance

g1<-ibox_emp_m%>%select(idea_date,month,ideano,`Location Name`,vertical)%>%group_by(month=format(as.Date(ibox_emp_m$idea_date),"%Y-%m"),`Location Name`,vertical)%>%filter(vertical=="GBS Finance")%>%summarise(count_ideas=n())
g1$month<-as.yearmon(g1$month)
g1$month<-as.Date(g1$month)
tail(g1)
p13<-ggplot(data=g1,aes(x=month,y=count_ideas),fill=col)+geom_line(color="dodgerblue2")+geom_point(color="black",size=2)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+geom_text(label=g1$count_ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)+theme(axis.text.x = element_text(angle=0, hjust = 1))+ggtitle("Monthly Trend of Ideas Generated - GBS Finance")+xlab("Months")+ylab("Numbers of Ideas Posted")+removeGrid()+facet_wrap(~`Location Name`)
p13

#Plot of Ibox Monthly trend - GBS Chennai and Pune - GBS PM
#g1<-ibox_emp_m%>%filter(vertical=="GBS PM")%>%select(month,ideano,`Location Name`)%>%group_by(month,`Location Name`)%>%summarise(count_ideas=n())
g1<-ibox_emp_m%>%select(idea_date,month,ideano,`Location Name`,vertical)%>%group_by(month=format(as.Date(ibox_emp_m$idea_date),"%Y-%m"),`Location Name`,vertical)%>%filter(vertical=="GBS PM")%>%summarise(count_ideas=n())
g1$month<-as.yearmon(g1$month)
g1$month<-as.Date(g1$month)
tail(g1)
p14<-ggplot(data=g1,aes(x=month,y=count_ideas),fill=col)+geom_line(color="dodgerblue2")+geom_point(color="black",size=2)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+geom_text(label=g1$count_ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)+theme(axis.text.x = element_text(angle=0, hjust = 1))+ggtitle("Monthly Trend of Ideas Generated - GBS PM")+xlab("Months")+ylab("Numbers of Ideas Posted")+removeGrid()+facet_wrap(~`Location Name`)
p14

##--------------------------------InProgress Ideas-------------------------------------------------

#Plot of Ibox Inprogress and Pending under each Management - GBS Chennai & GBS Pune
g2<-ibox_emp_m%>%select(vertical,Management_Chain_Level07,`Location Name`,new_status)%>%group_by(vertical,Management_Chain_Level07,`Location Name`,new_status)%>%summarise(count_ideas=n())%>%filter(!is.na(Management_Chain_Level07))
g2$Management_Chain_Level07<-gsub("[0-9]","",g2$Management_Chain_Level07)
g2$Management_Chain_Level07<-gsub("\\(","",g2$Management_Chain_Level07)
g2$Management_Chain_Level07<-gsub("\\)","",g2$Management_Chain_Level07)

g21<-g2%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Inprogress"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)

p2<-ggplot(data=g21,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity')+ggtitle("Inprogress Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=g21$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)


#Plot of Ibox Inprogress under each Management - GBS Chennai & GBS Pune -GBS Engineering
gg21<-g2%>%filter(vertical=="GBS Engineering")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Inprogress"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)
gg21_1<-gg21%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
p21<-ggplot(data=gg21_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p21 #Plot for GBS Engineering - Chennai
gg21_2<-gg21%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
p22<-ggplot(data=gg21_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p22 #Plot for GBS Engineering - Pune
p23<-p21+p22
p23 #Combined Plot for GBS Engineering

#Plot of Ibox Inprogress under each Management - GBS Chennai & GBS Pune - GBS Supply Chain
gg21<-g2%>%filter(vertical=="GBS Supply Chain")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Inprogress"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)
gg21_1<-gg21%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)

gg21_1
p21<-ggplot(data=gg21_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p21
gg21_2<-gg21%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
gg21_2
p22<-ggplot(data=gg21_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p22
p24<-p21+p22
p24 #Combined Plot for GBS Supply Chain

#Plot of Ibox Inprogress under each Management - GBS Chennai & GBS Pune - GBS Finance
gg21<-g2%>%filter(vertical=="GBS Finance")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Inprogress"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)
gg21_1<-gg21%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg21_1
p21<-ggplot(data=gg21_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p21
gg21_2<-gg21%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
gg21_2
p22<-ggplot(data=gg21_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p22
p25<-p21+p22
p25 #Combined Plot for GBS Finance


#Plot of Ibox Inprogress under each Management - GBS Chennai & GBS Pune - GBS Finance
gg21<-g2%>%filter(vertical=="GBS PM")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Inprogress"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)
gg21_1<-gg21%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg21_1
p21<-ggplot(data=gg21_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p21
gg21_2<-gg21%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
gg21_2
p22<-ggplot(data=gg21_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Inprogress_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg21_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p22
p26<-p21+p22
p26 #Combined Plot for GBS PM

##-----------------------------------Pending ideas---------------------------------------------------
g33<-g2%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Pending"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)
g33
p3<-ggplot(data=g33,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity')+ggtitle("Pending Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=g33$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p3

#Plot of Ibox pending under each Management - GBS Chennai & GBS Pune -GBS Engineering
gg31<-g2%>%filter(vertical=="GBS Engineering")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Pending"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)
gg31
gg31_1<-gg31%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg31_1
p31<-ggplot(data=gg31_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p31 #Plot for GBS Engineering - Chennai
gg31_2<-gg31%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
p32<-ggplot(data=gg31_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p32 #Plot for GBS Engineering - Pune
p33<-p31+p32
p33 #Combined Plot for GBS Engineering


#Plot of Ibox pending under each Management - GBS Chennai & GBS Pune -GBS Supply Chain
gg31<-g2%>%filter(vertical=="GBS Supply Chain")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Pending"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)
gg31
gg31_1<-gg31%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg31_1
p31<-ggplot(data=gg31_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p31 #Plot for GBS Supply Chain - Chennai
gg31_2<-gg31%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
p32<-ggplot(data=gg31_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p32 #Plot for GBS Supply Chain - Pune
p34<-p31+p32
p34 #Combined Plot for GBS Supply Chain


#Plot of Ibox pending under each Management - GBS Chennai & GBS Pune - GBS Finance
gg31<-g2%>%filter(vertical=="GBS Finance")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Pending"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)
gg31
gg31_1<-gg31%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg31_1
p31<-ggplot(data=gg31_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p31 #Plot for GBS Finance- Chennai
gg31_2<-gg31%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
p32<-ggplot(data=gg31_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p32 #Plot for GBS Finance - Pune
p35<-p31+p32
p35 #Combined Plot for GBS Finance

#Plot of Ibox pending under each Management - GBS Chennai & GBS Pune - GBS PM
gg31<-g2%>%filter(vertical=="GBS PM")%>%group_by(new_status,`Location Name`,Management_Chain_Level07)%>%filter(new_status %in% c("Pending"))%>%summarise(Total_Ideas=sum(count_ideas))%>%arrange(-Total_Ideas)%>%head(5)
gg31
gg31_1<-gg31%>%filter(`Location Name`=="GBS_Chennai")%>%top_n(5)
gg31_1
p31<-ggplot(data=gg31_1,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Chennai - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_1$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p31 #Plot for GBS PM - Chennai
gg31_2<-gg31%>%filter(`Location Name`=="GBS_Pune")%>%top_n(5)
p32<-ggplot(data=gg31_2,aes(x=reorder(Management_Chain_Level07,-Total_Ideas,sum),y=Total_Ideas,fill=Management_Chain_Level07))+geom_bar(stat='Identity',show.legend = FALSE)+ggtitle("GBS Pune - Pending_Ideas under each leader")+xlab("")+ylab("Numbers of Ideas")+removeGrid()+geom_text(label=gg31_2$Total_Ideas, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p32 #Plot for GBS PM - Pune
p36<-p31+p32
p36 #Combined Plot for GBS PM

##----------------------------------------Ibox-BE Team Status----------------------------------------
#Plot - BE team status of Ibox Ideas
g4<-ibox_emp_m%>%select(idea_date,month,new_status)%>%group_by(month=format(as.Date(idea_date),"%Y-%m"),new_status)%>%summarise(Total=n())
g4
g4$month<-as.yearmon(g4$month)
g4$month<-as.Date(g4$month)

p4<-ggplot(data=g4,aes(x=month,y=Total,fill=new_status))+geom_bar(stat='Identity',show.legend = FALSE)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+ggtitle("Status of Ideas - BE Team Status")+xlab("")+ylab("")+removeGrid()+facet_wrap(~new_status)+geom_text(label=g4$Total, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p4

#Plot of Ibox status - GBS Chennai & GBS Pune - GBS Engineering
g44_1<-ibox_emp_m%>%filter(vertical=="GBS Engineering")%>%select(idea_date,month,`Location Name`,new_status)%>%group_by(month=format(as.Date(idea_date),"%Y-%m"),`Location Name`,new_status)%>%summarise(Total=n())
g44_1$month<-as.yearmon(g44_1$month)
g44_1$month<-as.Date(g44_1$month)
g44_1
p41<-ggplot(data=g44_1,aes(x=month,y=Total,fill=new_status))+geom_bar(stat='Identity',show.legend = FALSE)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+ggtitle("Status of Ideas - BE Team Status - GBS Engineering")+xlab("")+ylab("")+removeGrid()+facet_wrap(~`Location Name`+new_status)+geom_text(label=g44_1$Total, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p41

#Plot of Ibox status - GBS Chennai & GBS Pune - GBS Supply Chain
g44_2<-ibox_emp_m%>%filter(vertical=="GBS Supply Chain")%>%select(idea_date,month,`Location Name`,new_status)%>%group_by(month=format(as.Date(idea_date),"%Y-%m"),`Location Name`,new_status)%>%summarise(Total=n())
g44_2$month<-as.yearmon(g44_2$month)
g44_2$month<-as.Date(g44_2$month)
g44_2
p42<-ggplot(data=g44_2,aes(x=month,y=Total,fill=new_status))+geom_bar(stat='Identity',show.legend = FALSE)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+ggtitle("Status of Ideas - BE Team Status - GBS Supply Chain")+xlab("")+ylab("")+removeGrid()+facet_wrap(~`Location Name`+new_status)+geom_text(label=g44_2$Total, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p42

#Plot of Ibox status - GBS Chennai & GBS Pune - GBS Finance
g44_3<-ibox_emp_m%>%filter(vertical=="GBS Finance")%>%select(idea_date,month,`Location Name`,new_status)%>%group_by(month=format(as.Date(idea_date),"%Y-%m"),`Location Name`,new_status)%>%summarise(Total=n())
g44_3$month<-as.yearmon(g44_3$month)
g44_3$month<-as.Date(g44_3$month)
p43<-ggplot(data=g44_3,aes(x=month,y=Total,fill=new_status))+geom_bar(stat='Identity',show.legend = FALSE)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+ggtitle("Status of Ideas - BE Team Status - GBS Finance")+xlab("")+ylab("")+removeGrid()+facet_wrap(~`Location Name`+new_status)+geom_text(label=g44_3$Total, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p43

#Plot of Ibox status - GBS Chennai & GBS Pune - GBS PM
g44_4<-ibox_emp_m%>%filter(vertical=="GBS PM")%>%select(idea_date,month,`Location Name`,new_status)%>%group_by(month=format(as.Date(idea_date),"%Y-%m"),`Location Name`,new_status)%>%summarise(Total=n())
g44_4$month<-as.yearmon(g44_4$month)
g44_4$month<-as.Date(g44_4$month)
g44_4
p44<-ggplot(data=g44_4,aes(x=month,y=Total,fill=new_status))+geom_bar(stat='Identity',show.legend = FALSE)+scale_x_date(date_breaks="1 month",date_labels = "%b'%y")+ggtitle("Status of Ideas - BE Team Status - GBS PM")+xlab("")+ylab("")+removeGrid()+facet_wrap(~`Location Name`+new_status)+geom_text(label=g44_4$Total, hjust= -0.5,vjust=-0.5,nudge_y = 0.2,check_overlap = TRUE,size=3)
p44


library(officer)
##----------------------------------Generating - GBS Engineering pptx----------------------------
my_pres<-read_pptx(path="C:/Users/gssaruba/Documents/Ibox_Template.pptx")
layout_summary(my_pres)

#creation of slide1 - Idea trend
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p11,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                  newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Engineering : Monthly Trend - Generation of Ideas",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                    newlabel = "", type = 'title'))
#creation of slide2 - BE Status
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p41,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Engineering : Idea Status",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                          newlabel = "", type = 'title'))
#creation of slide3 - Inprogress Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p23,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Engineering : Inprogress Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                  newlabel = "", type = 'title'))
#creation of slide4 - Pending Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p33,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))


my_pres <- ph_with(my_pres, value = "GBS Engineering : Pending Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                newlabel = "", type = 'title'))
#creation of slide5 - Thankyou
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres <- ph_with(my_pres, value = "Thank You ",location=ph_location_template(left = 7, top = 4, width = 8, height = 2,
                                                                                                             newlabel = "", type = 'body'))
print(my_pres, target = "GBS_Engineering_Ibox_dashboard.pptx")

##------------------------------Generating - GBS Supply Chain pptx-----------------------------------
my_pres<-read_pptx(path="C:/Users/gssaruba/Documents/Ibox_Template.pptx")
layout_summary(my_pres)

#creation of slide1 - Idea trend
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p12,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Supply Chain : Monthly Trend - Generation of Ideas",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                          newlabel = "", type = 'title'))
#creation of slide2 - BE Status
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p42,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Supply Chain : Idea Status",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                  newlabel = "", type = 'title'))
#creation of slide3 - Inprogress Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p24,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Supply Chain: Inprogress Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                newlabel = "", type = 'title'))
#creation of slide4 - Pending Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p34,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Supply Chain : Pending Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                             newlabel = "", type = 'title'))
#creation of slide5 - Thankyou
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres <- ph_with(my_pres, value = "Thank You ",location=ph_location_template(left = 7, top = 4, width = 8, height = 2,
                                                                               newlabel = "", type = 'body'))
print(my_pres, target = "GBS_SupplyChain_Ibox_dashboard.pptx")

##-------------------------------------Generating - GBS Finance pptx----------------------------------
my_pres<-read_pptx(path="C:/Users/gssaruba/Documents/Ibox_Template.pptx")
layout_summary(my_pres)

#creation of slide1 - Idea trend
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p13,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Finance : Monthly Trend - Generation of Ideas",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                           newlabel = "", type = 'title'))
#creation of slide2 - BE Status
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p43,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Finance : Idea Status",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                   newlabel = "", type = 'title'))
#creation of slide3 - Inprogress Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p25,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Finance: Inprogress Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                newlabel = "", type = 'title'))
#creation of slide4 - Pending Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p35,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS Finance : Pending Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                              newlabel = "", type = 'title'))
#creation of slide5 - Thankyou
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres <- ph_with(my_pres, value = "Thank You ",location=ph_location_template(left = 7, top = 4, width = 8, height = 2,
                                                                               newlabel = "", type = 'body'))
print(my_pres, target = "GBS_Finance_Ibox_dashboard.pptx")

##-----------------------------------------Generating - GBS PM pptx----------------------------------
my_pres<-read_pptx(path="C:/Users/gssaruba/Documents/Ibox_Template.pptx")
layout_summary(my_pres)

#creation of slide1 - Idea trend
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p14,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS PM : Monthly Trend - Generation of Ideas",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                           newlabel = "", type = 'title'))
#creation of slide2 - BE Status
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p44,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS PM : Idea Status",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                   newlabel = "", type = 'title'))
#creation of slide3 - Inprogress Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p26,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS PM: Inprogress Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                                newlabel = "", type = 'title'))
#creation of slide4 - Pending Ideas
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres<-ph_with(x=my_pres,value=p36,location=ph_location_template(left = 1, top = 2, width = 16, height = 7,
                                                                   newlabel = "", type = NULL, id = 1))
my_pres <- ph_with(my_pres, value = "GBS PM : Pending Ideas Summary ",location=ph_location_template(left = 1, top = 0, width = 15, height = 2,
                                                                                                              newlabel = "", type = 'title'))
#creation of slide5 - Thankyou
my_pres<-add_slide(my_pres,layout="Blank Slide",master="1_Flex_template_3.0")
my_pres <- ph_with(my_pres, value = "Thank You ",location=ph_location_template(left = 7, top = 4, width = 8, height = 2,
                                                                               newlabel = "", type = 'body'))
print(my_pres, target = "GBS_PM_Ibox_dashboard.pptx")




##------------------------------------Emailing GBS Engineering ---------------------------------------- 
library(RDCOMClient)
Outlook <- COMCreate("Outlook.Application")
Email = Outlook$CreateItem(0)
Email[["sentonbehalfofname"]] = "gbschennaibe.team@flex.com"
#Email[["to"]] = "prasannakumar.anbazhagan@flex.com;sujatha.rathishm@flex.com"
Email[["cc"]] = ""
Email[["bcc"]] = "arunrajesh.balakrishnan@flex.com"
Email[["subject"]] = "Ibox Visual dashboard"
Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/ibox_data.csv")
Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/GBS_Engineering_Ibox_dashboard.pptx")
Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/GBS_SupplyChain_Ibox_dashboard.pptx")
Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/GBS_Finance_Ibox_dashboard.pptx")
Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/GBS_PM_Ibox_dashboard.pptx")
Email[["htmlbody"]] = "<p>Hi Team,
<br/>
<br>Please find attached Ibox visual dashboard and the corresponding source data for the same.
<br/>
<br/>
<br><strong> <i>Note: This is an autogenerated visual dashbaord. Please contact your respective SPOC for any clarification</i></strong>
<br/>
<br/>
<br>Thank You
<br/>
<br>Regards<br/>
<br><i>Business Excellence Team</i></p>"
Email$Send()
#Close Outlook, clear the message
rm(Outlook, Email)

##------------------------------------Emailing GBS Supply Chain------------------------------------------------------------------
# library(RDCOMClient)
# Outlook <- COMCreate("Outlook.Application")
# Email = Outlook$CreateItem(0)
# Email[["sentonbehalfofname"]] = "gbschennaibe.team@flex.com"
# Email[["to"]] = "ramesh.bandaru@flex.com;prashant.kamal@flex.com;pavithra.kannan@flex.com;nandhini.ramachandran@flex.com;baranitharan.gopalakrishnan@flex.com;bharath.ragothaman@flex.com"
# Email[["cc"]] = "vishnu.cheche@flex.com;sujatha.rathishm@flex.com;jayakumar.s@flex.com;arunrajesh.balakrishnan@flex.com"
# Email[["bcc"]] = ""
# Email[["subject"]] = "GBS Supply Chain Ibox dashboard"
# Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/ibox_data.csv")
# Email[["Attachments"]]$Add("C:/Users/gssaruba/Documents/GBS_SupplyChain_Ibox_dashboard.pptx")
# Email[["htmlbody"]] = "<p>Hi Team,
# <br/>
# <br>Please find attached Ibox visual dashboard and the corresponding source data for the same.
# <br/>
# <br/>
# <br><strong> <i>Note: This is an autogenerated visual dashbaord. Please contact your respective SPOC for any clarification</i></strong>
# <br/>
# <br/>
# <br>Thank You
# <br/>
# <br>Regards<br/>
# <br><i>Business Excellence Team</i></p>"
# Email$Send()
# #Close Outlook, clear the message
# rm(Outlook, Email)

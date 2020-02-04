library(readxl)
library(writexl)
library(tidyverse)
library(rio)

setwd("//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS INDIA")

#-------------------Set working directory to GBS INDIA-------------------------------------
setwd("//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS INDIA/December 2019")
file_list <- list.files(path="//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS INDIA/December 2019")
file_list<-file_list[6:10]
file_list

# Declaring empty dataframes
dataset_rpa_india<-data.frame()
dataset_macro_it_india<-data.frame()
dataset_report_india<-data.frame()

#Executing for loop for reading RPA dataset
for (i in 1:length(file_list))
  {
  temp_data1<- read_excel(file_list[i],sheet = 1,skip = 1)
  temp_data11<-temp_data1%>%mutate_all(as.character)
  dataset_rpa_india<-rbind(dataset_rpa_india,temp_data11)
  }


#Executing for loop for reading IT Automation and Macro dataset
for (i in 1:length(file_list))
  {
 temp_data2<- read_excel(file_list[i],sheet =2,skip = 1)
 colnames(temp_data2)
 temp_data22<-temp_data2%>%mutate_all(as.character)
 dataset_macro_it_india<-rbind(dataset_macro_it_india,temp_data22)
}


#Executing for loop for reading Report dataset
for (i in 1:length(file_list))
  {
  temp_data3<- read_excel(file_list[i],sheet =3,skip=1)
  colnames(temp_data3)
  temp_data33<-temp_data3%>%mutate_all(as.character)
  dataset_report_india<-rbind(dataset_report_india,temp_data33)
 }

ssample<-rbind(dataset_rpa_india,dataset_macro_it_india)
getwd()
#write.csv(ssample,"//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects/PowerBiSource/Oct2019/India_rpa_and_macro_automation.csv")
#write.csv(dataset_report_india,"//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects/PowerBiSource/Oct2019/India_report.csv")
export(ssample,"India_rpa_and_macro_automation.xlsx",format="xlsx")
export(dataset_report_india,"India_report.xlsx",format="xlsx")


#----------------Set working directory to GBS GAUD--------------------------------------
setwd("//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS GUAD/December 2019/")

file_list <- list.files(path="//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS GUAD/December 2019")
file_list<-file_list[2]
file_list

# Declaring empty dataframes
dataset_rpa_guad<-data.frame()
dataset_macro_it_guad<-data.frame()
dataset_report_guad<-data.frame()

getwd()

#Executing for loop for reading RPA dataset
for (i in 1:length(file_list))
{
  temp_data1<- read_excel(file_list[i],sheet = 1,skip=1)
  temp_data11<-temp_data1%>%mutate_all(as.character)
  dataset_rpa_guad<-rbind(dataset_rpa_guad,temp_data11)
}
#Executing for loop for reading IT Automation and Macro dataset
for (i in 1:length(file_list))
{
  temp_data2<- read_excel(file_list[i],sheet =2,skip = 1)
  temp_data22<-temp_data2%>%mutate_all(as.character)
  dataset_macro_it_guad<-rbind(dataset_macro_it_guad,temp_data22)
}
#Executing for loop for reading Report dataset
for (i in 1:length(file_list))
{
  temp_data3<- read_excel(file_list[i],sheet =3,skip=1)
  temp_data33<-temp_data3%>%mutate_all(as.character)
  dataset_report_guad<-rbind(dataset_report_guad,temp_data33)
}

#----------------Set working directory to GBS SHENZHEN--------------------------------------
setwd("//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS SHENZHEN/December 2019")
file_list <- list.files(path="//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects and Productivity/GBS SHENZHEN/December 2019")
file_list <-file_list[1]
file_list

# Declaring empty dataframes
dataset_rpa_shen<-data.frame()
dataset_macro_it_shen<-data.frame()
dataset_report_shen<-data.frame()

#Executing for loop for reading RPA dataset
for (i in 1:length(file_list))
{
  temp_data1<- read_excel(file_list[i],sheet = 1,skip = 1)
  temp_data11<-temp_data1%>%mutate_all(as.character)
  dataset_rpa_shen<-rbind(dataset_rpa_shen,temp_data11)
}
#Executing for loop for reading IT Automation and Macro dataset
for (i in 1:length(file_list))
{
  temp_data2<- read_excel(file_list[i],sheet =2,skip=1)
  temp_data22<-temp_data2%>%mutate_all(as.character)
  dataset_macro_it_shen<-rbind(dataset_macro_it_shen,temp_data22)
}
#Executing for loop for reading Report dataset
for (i in 1:length(file_list))
{
  temp_data3<- read_excel(file_list[i],sheet =3,skip=1)
  temp_data33<-temp_data3%>%mutate_all(as.character)
  dataset_report_shen<-rbind(dataset_report_shen,temp_data33)
}


con_rpa<-rbind(dataset_rpa_india,dataset_rpa_shen,dataset_rpa_guad)
con_macro<-rbind(dataset_macro_it_india,dataset_macro_it_shen,dataset_macro_it_guad)
con_report<-rbind(dataset_report_india,dataset_report_shen,dataset_report_guad)
con_rpa_macro<-rbind(con_rpa,con_macro)


#write.csv(con_rpa_macro,"//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects/PowerBiSource/Oct2019/rpa_and_macro_automation.csv")
#write.csv(con_report,"//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects/PowerBiSource/Oct2019/consolidated_report.csv")
setwd("//flextronics365.sharepoint.com@SSL/DavWWWRoot/sites/gbs_productivity/GBS Projects/PowerBiSource/Dec2019")

export(con_rpa_macro,"rpa_and_macro_automation.xlsx",format = "xlsx")
export(con_report,"consolidated_report.xlsx",format="xlsx")


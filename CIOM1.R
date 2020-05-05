library(readxl)
library(stringr)
library(ggplot2)
library(dplyr)
library(rlist)
library(xlsx)
library(DescTools)
library(binom)
#how many meet each level
getNumLevelStat = function(factor)
{
  lev = levels(factor)
  d = data.frame()
  level = c()
  value = c()
  for (k in lev) {
    l = k
    v = 0
    for (i in factor) {
      if (i == l)
        v = v + 1
    }
    level = c(level,l)
    value = c(value,as.integer(v))
  }
  d = data.frame(level,value)
  return(d)
}
#get mark
getRating = function(x)
{
  mark = 0
  weight_Time = ifelse(x$When == "Обучались в этом учебном году",weight_Now,weight_Early)
  rate = weight_Time*(x$Program*weight_Prog + x$Culture*weight_Cult + x$Organization*weight_Org+ x$Provision*weight_MTO)
  #if we decide to use polynomial distribution
  # if (rate < 0.5*MaxMark)
  #   mark = 1
  # else if (rate >= 0.5*MaxMark && rate < 0.7*MaxMark)
  #   mark = 2
  # else if (rate >= 0.7*MaxMark && rate < 0.8*MaxMark)
  #   mark = 3
  # else if (rate >= 0.8*MaxMark && rate < 0.9*MaxMark)
  #   mark = 4
  # else if (rate >= 0.9*MaxMark)
  #   mark = 5
  #if we decide to use binomial distribution
  mark = ifelse(rate < 0.7 * MaxMark,0,1)
  return(mark)
}

#read and prepare data
setwd("C:\\Users\\SmirnygaTotoshka\\Desktop\\RWorkspace")
allTable = read_excel("YourVoice.xlsx",sheet = 2)

allTable$Cafedra = as.factor(allTable$Cafedra)
allTable$Time = as.Date.character(allTable$Time)
allTable$Cafedra = str_replace_all(allTable$Cafedra,":","_")
departments = levels(allTable$Cafedra)

#weight coefficients - constant!
weight_Prog <- 0.7
weight_Cult <- 0.4
weight_Org <- 1
weight_MTO <- 0
weight_Now = 1
weight_Early = 1
MaxMark = 5*(weight_Prog + weight_Cult + weight_Org + weight_MTO)

#all criteria to one mark - like/dislike
marks = c()

for(i in 1:nrow(allTable))
{
  marks = c(marks,getRating(allTable[i,]))
}
allTable$Mark = marks
#get weeks. its needed to calculate moving average
allTable$Week = as.integer(strftime(allTable$Time,format = "%V"))
allTable$Year = as.integer(strftime(allTable$Time,format = "%Y"))
allTable$ShiftedWeek = 0
w = grep("^Week$", colnames(allTable))
y = grep("^Year$", colnames(allTable))
sw = grep("^ShiftedWeek$", colnames(allTable))
c = 0
#shift week
for (i in 2:nrow(allTable)) {
  if(allTable[i,w] != allTable[i-1,w] || allTable[i,y] != allTable[i-1,y])
    c = c + 1
  allTable[i,sw] = c
}
onDepartaments = split(allTable, f = allTable$Cafedra)
#calculate moving average for department on weeks
#there are marks accumalation 
#df - data frame with department marks
#outputPlot - path to save plot
moving.average.for.department = function(df,outputPlot)
{
  moving.average = data.frame()
  for (i in unique(df$ShiftedWeek)) {
    tmp = subset(df,ShiftedWeek <= i)
    avg_Prog = mean(tmp$Program)
    avg_Cult = mean(tmp$Culture)
    avg_Org = mean(tmp$Organization)
    avg_MTO = mean(tmp$Provision)
    tmp = data.frame(
      avgProgram = avg_Prog,
      avgCulture = avg_Cult,
      avgOrganization = avg_Org,
      avgProvision = avg_MTO,
      ShiftedWeek = i)
    moving.average = rbind.data.frame(moving.average,tmp)
  }
  p = ggplot()+
    geom_line(aes(x = ShiftedWeek,y = avgProgram,color='black'),moving.average)+geom_point(aes(x = ShiftedWeek,y = avgProgram,color='black'),moving.average)+
    geom_line(aes(x = ShiftedWeek,y = avgCulture,color = "red"),moving.average)+geom_point(aes(x = ShiftedWeek,y = avgCulture,color = "red"),moving.average)+
    geom_line(aes(x = ShiftedWeek,y = avgOrganization,color = "green"),moving.average)+geom_point(aes(x = ShiftedWeek,y = avgOrganization,color = "green"),moving.average)+
    geom_line(aes(x = ShiftedWeek,y = avgProvision,color = "blue"),moving.average)+geom_point(aes(x = ShiftedWeek,y = avgProvision,color = "blue"),moving.average)+
    labs(x="Week",y="Mark moving average",title=moving.average$Cafedra[1])+
    scale_colour_manual(name = 'Legend',
                        values =c('black'='black','red'='red','green' = 'green','blue'='blue'), labels = c('Subject program','Lecturer culture','Organization','Provision'))+
    ylim(1,5)
  ggsave(filename = outputPlot,plot = p, width = 508, height = 285.75, units = 'mm', dpi = "retina", device = "png")
  return(moving.average)
}
#for departments where num of answers > 50
for (i in onDepartaments) {
  if(nrow(i) >= 50)
    moving.average.for.department(i,paste(getwd(),"/more50/",i$Cafedra[1],".png",sep = ""))
}

all = data.frame()
#type_marks = c(1,2,3,4,5)
#Calculate rating of the departments how ci of binom distribution(like/dislike)
for(i in onDepartaments)
{
    k = nrow(i)
    avg_Prog = mean(i$Program)
    avg_Cult = mean(i$Culture)
    avg_Org = mean(i$Organization)
    avg_MTO = mean(i$Provision)

    # numMarks = getNumLevelStat(factor(i$Mark,levels = type_marks))
    # ci = as.data.frame(MultinomCI(numMarks$value,conf.level = 0.90,method = "wilson"))
    # low_rate = sum(ci$lwr.ci * type_marks) / 5 * 100
    # high_rate = sum(ci$upr.ci * type_marks) / 5 * 100
    # #rating cannot be more 100
    # high_rate = ifelse(high_rate > 100,100,high_rate)
    low_rate = round(binom.confint(sum(i$Mark), length(i$Mark),conf.level = 0.90, methods = "wilson")$lower,4)*100
    high_rate = round(binom.confint(sum(i$Mark), length(i$Mark),conf.level = 0.90, methods = "wilson")$upper,4)*100
    delta_rate = high_rate - low_rate

    df = data.frame(dep = i$Cafedra[1],answers = k,mean_Prog = avg_Prog,
                    mean_Cult = avg_Cult,mean_Org = avg_Org,mean_MTO = avg_MTO,
                    low_rate = low_rate,high_rate = high_rate,delta_rate = delta_rate)
    all = rbind.data.frame(all,df)
}

#get moving average for top 10 departments
o = order(all$low_rate,decreasing = T)
o = o[1:10]
for(i in o)
{
  df = onDepartaments[[i]]
  moving.average.for.department(df,paste(getwd(),"/top10/",df$Cafedra[1],".png",sep = ""))
}

#get order each depatment on each position
ordered_all = all
colnames(ordered_all) = c("Dep","Answers","placeProg","placeCult","placeOrg","placeMTO","low_rate","high_rate","delta_rate")
ordProg = order(all$mean_Prog,decreasing = T)
ordCult = order(all$mean_Cult,decreasing = T)
ordOrg = order(all$mean_Org,decreasing = T)
ordMTO = order(all$mean_MTO,decreasing = T)
iProg = grep("^placeProg$", colnames(ordered_all))
iCult = grep("^placeCult$", colnames(ordered_all))
iOrg = grep("^placeOrg$", colnames(ordered_all))
iMTO = grep("^placeMTO$", colnames(ordered_all))
iAns = grep("^Answers$",colnames(ordered_all))
for (i in 1:nrow(ordered_all)) {
  if (ordered_all[i,iAns] < 15)
  {
    ordered_all[i,iProg] = ordered_all[i,iCult] = ordered_all[i,iOrg] = ordered_all[i,iMTO] = "Not enough data"
  }
  else
  {
    nProg=match(i,ordProg)
    nCult=match(i,ordCult)
    nOrg=match(i,ordOrg)
    nMTO=match(i,ordMTO)
    ordered_all[i,iProg] = paste(nProg,"/",nrow(ordered_all),sep="")
    ordered_all[i,iCult] = paste(nCult,"/",nrow(ordered_all),sep="")
    ordered_all[i,iOrg] = paste(nOrg,"/",nrow(ordered_all),sep="")
    ordered_all[i,iMTO] = paste(nMTO,"/",nrow(ordered_all),sep="")
    
  }
}

#save final results. beauty! on Russian(merge all and ordered_all without duplicates)
final = data.frame(all$dep,all$answers,all$mean_Prog,ordered_all$placeProg,
                   all$mean_Cult,ordered_all$placeCult,all$mean_Org,ordered_all$placeOrg,
                   all$mean_MTO,ordered_all$placeMTO,all$low_rate,all$high_rate,all$delta_rate)
colnames(final)=c("Кафедра","Количество ответов","ПП - среднее","ПП - положение",
                  "КП - среднее","КП - положение","ОУП - среднее","ОУП - положение",
                  "МТО - среднее","МТО - положение","Нижняя граница рейтинга",
                  "Верхняя граница рейтинга","Интервал реализации")

write.xlsx(final,"FinalYourVoice.xlsx",sheetName = "all",append = F,row.names = F)



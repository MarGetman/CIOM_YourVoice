library(readxl)
library(stringr)
library(ggplot2)
library(dplyr)
library(rlist)
library(xlsx)
library(DescTools)
library(binom)
setwd("C:\\Users\\SmirnygaTotoshka\\Desktop\\RWorkspace")
allTable = read_excel("YourVoice.xlsx",sheet = 2)

allTable$Cafedra = as.factor(allTable$Cafedra)
allTable$Time = as.Date.character(allTable$Time)

departments = levels(allTable$Cafedra)

#weight coefficients
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

onDepartaments = split(allTable, f = allTable$Cafedra)

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
#Test code is commented
# class(df$Mark)
# factor(df$Mark,levels = c(0,1))
# nums = getNumLevelStat(factor(df$Mark,levels = c(1,2,3,4,5)))
# MultinomCI(nums$value,conf.level = 0.90,method = "wilson",sides = "two.sided")
# binom.confint(sum(df$Mark), length(df$Mark),conf.level = 0.90, methods = "wilson")
all = data.frame()
type_marks = c(1,2,3,4,5)
for(i in onDepartaments)
{
    k = nrow(i)
    avg_Prog = mean(i$Program)
    avg_Cult = mean(i$Culture)
    avg_Org = mean(i$Organization)
    avg_MTO = mean(i$Provision)
    
    numMarks = getNumLevelStat(factor(i$Mark,levels = type_marks))
    ci = as.data.frame(MultinomCI(nums$value,conf.level = 0.90,method = "wilson"))
    low_rate = sum(ci$lwr.ci * type_marks) / 5 * 100
    high_rate = sum(ci$upr.ci * type_marks) / 5 * 100
    #low_rate = round(binom.confint(sum(i$Mark), length(i$Mark),conf.level = 0.90, methods = "wilson")$lower,4)*100
    #high_rate = round(binom.confint(sum(i$Mark), length(i$Mark),conf.level = 0.90, methods = "wilson")$upper,4)*100
    delta_rate = high_rate - low_rate
    
    df = data.frame(dep = i$Cafedra[1],answers = k,mean_Prog = avg_Prog,
                    mean_Cult = avg_Cult,mean_Org = avg_Org,mean_MTO = avg_MTO,
                    low_rate = low_rate,high_rate = high_rate,delta_rate = delta_rate)
    all = rbind.data.frame(all,df)
}
#write.xlsx(all,"allRating.xlsx",sheetName = "AllRating",append = F,row.names = F)

#n = subset(n,nrow >= 20)

#get mark
getRating = function(x)
{
  mark = 0
  weight_Time = ifelse(x$When == "Обучались в этом учебном году",weight_Now,weight_Early)
  rate = weight_Time*(x$Program*weight_Prog + x$Culture*weight_Cult + x$Organization*weight_Org+ x$Provision*weight_MTO)
  #if we decide to use polynomial distribution
  if (rate < 0.5*MaxMark)
    mark = 1
  else if (rate >= 0.5*MaxMark && rate < 0.7*MaxMark)
    mark = 2
  else if (rate >= 0.7*MaxMark && rate < 0.8*MaxMark)
    mark = 3
  else if (rate >= 0.8*MaxMark && rate < 0.9*MaxMark)
    mark = 4
  else if (rate >= 0.9*MaxMark)
    mark = 5
  #if we decide to use binomial distribution
  #mark = ifelse(rate < 0.7 * MaxMark,0,1)
  return(mark)
}



#All code below NOT RUN!
# mon = months(marks[[departments[1]]]$Time)
# itog = list()
# c = 1
# for (i in marks) {
#   #mon = months(i[[departments[departments = i$Cafedra]]]$Time)
#   mon = months(i$Time)
#   i$Month = mon
#   depOnMon = split(i, f = i$Month)
#   df= data.frame()
#   for (j in depOnMon) {
#     m = j$Month[1]
#     avg_Prog = mean(j$Program)
#     avg_Cult = mean(j$Culture)
#     avg_Org = mean(j$Organization)
#     avg_MTO = mean(j$Provision)
#     tmp = data.frame(Month = m,
#                      avgProgram = avg_Prog,
#                      avgCulture = avg_Cult,
#                      avgOrganization = avg_Org,
#                      avgProvision = avg_MTO)
#     df = rbind.data.frame(df,tmp)
#   }
#   itog[[departments[c]]] = df
#   c = c+1
# }
# # write to excel - Not run
# for (i in departments) {
#   name = str_replace_all(i,":","-")
#   
#   write.xlsx(itog[[i]],paste("Results/ResultsYourVoice",name,".xlsx"),sheetName = name,append = F,row.names = F)
# }

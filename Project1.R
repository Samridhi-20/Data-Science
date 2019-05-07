# Clearing the environment
rm(list=ls(all=T))
# Setting working directory
setwd("E:/Edwisor/Project1")
getwd()

#Load Libraries
x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
      "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees')

#install.packages(x)
lapply(x, require, character.only = TRUE)

install.packages("xlsx")
library(xlsx)

## Reading the data
df_absent_ds = read.xlsx('Absenteeism_at_work_Project.xls', sheetIndex = 1)


#----------------------------------------------------Exploratory Data Analysis------------------------------------------------------
# Shape of the data
dim(df_absent_ds)
# Viewing data
View(df_absent_ds)
# Structure of the data
str(df_absent_ds)
# Variable namesof the data
colnames(df_absent_ds)
#summary of the data set
summary(df_absent_ds$ID)
#
class(df_absent_ds)


absent_ds_continuous = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Weight', 'Height', 
                    'Body.mass.index', 'Absenteeism.time.in.hours')

absent_ds_categorical = c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')



#------------------------------------Missing Values Analysis---------------------------------------------------#
#Creating dataframe with missing values present in each variable
missing_val = data.frame(apply(df_absent_ds,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"

#Calculating percentage missing value
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df_absent_ds)) * 100

# Sorting missing_val in Descending order
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL

# Reordering columns
missing_val = missing_val[,c(2,1)]

# Saving output result into csv file
write.csv(missing_val, "Missing_perc_R.csv", row.names = F)


#-----------------------Computing Missing values-----------------
# # Plotting the ggplot for missing values
 ggplot(data = missing_val[1:18,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
 geom_bar(stat = "identity",fill = "orange")+xlab("Variables")+
 ggtitle("Variables Vs Missing Percent Plot") + theme_bw()

#selecting a known variable and replacing its value as NA to find the best method for Missing value computation
#Performing the anaylsis over BMI column since it has higghest percent of missing values

df_absent_ds$Body.mass.index[10]
#Actual Value = 29

df_absent_ds$Body.mass.index[10]= NA

#Mean Method
df_absent_ds$Body.mass.index[is.na(df_absent_ds$Body.mass.index)] = mean(df_absent_ds$Body.mass.index, na.rm = T)

df_absent_ds$Body.mass.index[10]
## value by Mean=26.68079

#Reloading the dataset and again making the df_absent_ds$Body.mass.index[9] value as NA


#Median Method
df_absent_ds$Body.mass.index[is.na(df_absent_ds$Body.mass.index)]=median(df_absent_ds$Body.mass.index, na.rm = T)

df_absent_ds$Body.mass.index[10]
## value by Median=25

#packages and libraries req for KNN method
install.packages("DMwR")

library(DMwR)

##KNN function is also available under VIM package and library with name - 'kNN'. However the 
#disadvantage is that by using that i was getting my variables count as double as it was inserting one extra column as 'col_imp'
#for each column

# kNN Imputation
df_absent_ds = knnImputation(df_absent_ds, k = 5)

df_absent_ds$Body.mass.index[10]
## value by KNN=29

# Checking for missing value
sum(is.na(df_absent_ds))

#------------------------Histograms and Bar Plot Visualizations--------------
for (i in absent_ds_continuous){
  hist(df_absent_ds[,i], col="yellow", prob=TRUE, main="Histogram",xlab=i, ylab="Count")
  lines(density(df_absent_ds[,i]), lwd=2)
}


dev.new()

bar1 = ggplot(data = df_absent_ds, aes(x = Social.drinker, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity", fill="orange") + 
  ggtitle("Count of Reason for drinking") + theme_bw()

bar2 = ggplot(data = df_absent_ds, aes(x = Disciplinary.failure, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Dicipline Action") + theme_bw()

bar3 = ggplot(data = df_absent_ds, aes(x = ID, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for ID") + theme_bw()

bar4 = ggplot(data = df_absent_ds, aes(x = Month.of.absence, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Month of absence") + theme_bw()

bar5 = ggplot(data = df_absent_ds, aes(x = Day.of.the.week, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Day of week") + theme_bw()

bar6 = ggplot(data = df_absent_ds, aes(x = Seasons, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Season") + theme_bw()

bar7 = ggplot(data = df_absent_ds, aes(x = Social.smoker, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Smoking") + theme_bw()

bar8 = ggplot(data = df_absent_ds, aes(x = Son, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Son") + theme_bw()

bar9 = ggplot(data = df_absent_ds, aes(x = Pet, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Pet") + theme_bw()

bar10 = ggplot(data = df_absent_ds, aes(x = Education, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Education") + theme_bw()

bar11 = ggplot(data = df_absent_ds, aes(x = Reason.for.absence, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Education") + theme_bw()

gridExtra::grid.arrange(bar1,bar2,bar3,bar4,ncol=2)

gridExtra::grid.arrange(bar5,bar6,bar7,bar11,ncol=2)

gridExtra::grid.arrange(bar8,bar9,ncol=2)

gridExtra::grid.arrange(bar10,ncol=2)



#-------------------------------------Outlier Analysis-------------------------------------#
# BoxPlots - Distribution and Outlier Check

#dev.new()
# Boxplot for continuous variables
for (i in 1:length(absent_ds_continuous))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (absent_ds_continuous[i]), x = "Absenteeism.time.in.hours"), data = subset(df_absent_ds))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="black", fill = "orange" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=absent_ds_continuous[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of absenteeism for",absent_ds_continuous[i])))
}

# ## Plotting plots together
gridExtra::grid.arrange(gn1,gn2,ncol=2)
gridExtra::grid.arrange(gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,ncol=2)
gridExtra::grid.arrange(gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)


# #Remove outliers using boxplot method

# #loop to remove from all variables
#for(i in absent_ds_continuous)
#{
 # print(i)
  #val = df_absent_ds[,i][df_absent_ds[,i] %in% boxplot.stats(df_absent_ds[,i])$out]
  #print(length(val))
  #df_absent_ds = df_absent_ds[which(!df_absent_ds[,i] %in% val),]
#}

#Replace all outliers with NA and impute
for(i in absent_ds_continuous)
{
  val = df_absent_ds[,i][df_absent_ds[,i] %in% boxplot.stats(df_absent_ds[,i])$out]
  #print(length(val))
  df_absent_ds[,i][df_absent_ds[,i] %in% val] = NA
}

# Imputing missing values
df_absent_ds = knnImputation(df_absent_ds,k=5)


#-----------------------------------Feature Selection------------------------------------------#

library(corrgram)

## Correlation Plot 
corrgram(df_absent_ds[,absent_ds_continuous], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

## ANOVA test for Categprical variable
summary(aov(formula = Absenteeism.time.in.hours~ID,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Reason.for.absence,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Month.of.absence,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Day.of.the.week,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Seasons,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Disciplinary.failure,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Education,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Social.drinker,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Social.smoker,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Son,data = df_absent_ds))
summary(aov(formula = Absenteeism.time.in.hours~Pet,data = df_absent_ds))


## Dimension Reduction
df_absent_ds = subset(df_absent_ds, select = -c(Body.mass.index))


#--------------------------------Feature Scaling--------------------------------#
#Normality check
hist(df_absent_ds$Absenteeism.time.in.hours)

# Updating the continuous and catagorical variable
absent_ds_continuous = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Height', 'Weight')

absent_ds_categorical = c('ID','Reason.for.absence','Disciplinary.failure', 
                     'Social.drinker', 'Son', 'Pet', 'Month.of.absence', 'Day.of.the.week', 'Seasons',
                     'Education', 'Social.smoker')


# Normalization
for(i in absent_ds_continuous)
{
  print(i)
  df_absent_ds[,i] = (df_absent_ds[,i] - min(df_absent_ds[,i]))/(max(df_absent_ds[,i])-min(df_absent_ds[,i]))
}


#------------------------------------------Model Development--------------------------------------------#


#Divide data into train and test using stratified sampling method
set.seed(123)
train.index = sample(1:nrow(df_absent_ds), 0.8 * nrow(df_absent_ds))
train = df_absent_ds[ train.index,]
test  = df_absent_ds[-train.index,]

##Decision tree for classification
#Develop Model on training data

library(rpart)

df_dtree=rpart(Absenteeism.time.in.hours ~., data = train, method = "anova")

dev.new()
plot(df_dtree)
text(df_dtree)
summary(df_dtree)
printcp(df_dtree)



#write rules into disk
write(capture.output(summary(df_dtree)), "Rules.txt")

#Lets predict for test data
pred_test_DT = predict(df_dtree,test[,-20])

#Lets predict for train data
pred_train_DT = predict(df_dtree,train[,-20])

install.packages("caret")
library(caret)

#Since value to be predicted is of time series type hence using postresample 

# For training data 
print(postResample(pred = pred_train_DT, obs = train[,20]))
#     RMSE  Rsquared       MAE 
#2.5255970 0.4818267 1.7699633 

# For testing data 
print(postResample(pred = pred_test_DT, obs = test[,20]))
#     RMSE  Rsquared       MAE 
#2.4175810 0.3997683 1.5640250


#-----------------Linear Regression-----------------------


#check multicollearity
install.packages("usdm")
library(usdm)
vif(df_absent_ds[,-20])

vifcor(df_absent_ds[,-20], th = 0.9)

#run regression model
LR_model = lm( Absenteeism.time.in.hours ~., data = train)

#Summary of the model
summary(LR_model)


#Lets predict for test data
pred_test_LR = predict(LR_model,test[,-20])

#Lets predict for train data
pred_train_LR = predict(LR_model,train[,-20])

#Predict
predictions_LR = predict(LR_model, test[,1:19])

# For training data 
print(postResample(pred = pred_train_LR, obs = train[,20]))
#         RMSE  Rsquared       MAE 
#2.9313143 0.3019738 2.1895013   

# For testing data 
print(postResample(pred = pred_test_LR, obs = test[,20]))
#     RMSE  Rsquared       MAE 
#2.583734 0.326454 1.786583   

#------------------RANDOM FOREST------------------

install.packages("randomForest")
library(randomForest)
#RUN RANDOM FOREST
RF_model = randomForest(Absenteeism.time.in.hours ~., train, importance = TRUE, ntree = 500)

#Summary of the model
summary(RF_model)


#Lets predict for test data
pred_test_RF = predict(RF_model,test[,-20])

#Lets predict for train data
pred_train_RF = predict(RF_model,train[,-20])

#Predict
predictions_RF = predict(RF_model, test[,1:19])

# For training data 
print(postResample(pred = pred_train_RF, obs = train[,20]))
#         RMSE  Rsquared       MAE 
#1.4625432 0.8576811 1.0402535  

# For testing data 
print(postResample(pred = pred_test_RF, obs = test[,20]))
#     RMSE  Rsquared       MAE 
#2.3228404 0.4485493 1.5629652


#------------------to predict for 2011 trends------------
dev.new()

bar4 = ggplot(data = df_absent_ds, aes(x = Month.of.absence, y = Absenteeism.time.in.hours)) + geom_bar(stat="identity",fill ="orange") +
  ggtitle("Count of Reason for Month of absence") + theme_bw()

gridExtra::grid.arrange(bar4,ncol=2)

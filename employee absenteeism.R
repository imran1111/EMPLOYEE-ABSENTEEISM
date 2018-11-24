# Clearing the environment
rm(list=ls(all=T))
# Setting working directory
setwd("E:/R Learning")

# Loading libraries
x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
      "MASS", "rpart","mlr", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees','fastDummies','xlsx')

lapply(x, require, character.only = TRUE)
rm(x)

## Reading the data
absnt = read.xlsx('Absenteeism_at_work_Project.xls', sheetIndex = 1)


#----------------------------------------------------Exploratory Data Analysis------------------------------------------------------
# Shape of the data
dim(absnt)
# View(absnt)
# Structure of the data
str(absnt)
# Variable namesof the data
colnames(absnt)


# From the above EDA and problem statement categorising data in 2 category "continuous" and "catagorical"
continuous_vars = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Weight', 'Height', 
                    'Body.mass.index', 'Absenteeism.time.in.hours')

catagorical_vars = c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')
absnt$ID=as.factor(absnt$ID)
absnt$Reason.for.absence=as.factor(absnt$Reason.for.absence)
absnt$Month.of.absence=as.factor(absnt$Month.of.absence)
absnt$Day.of.the.week=as.factor(absnt$Day.of.the.week)
absnt$Seasons=as.factor(absnt$Seasons)
absnt$Disciplinary.failure=as.factor(absnt$Disciplinary.failure)
absnt$Education=as.factor(absnt$Education)
absnt$Social.drinker=as.factor(absnt$Social.drinker)
absnt$Social.smoker=as.factor(absnt$Social.smoker)
absnt$Pet=as.factor(absnt$Pet)
absnt$Son=as.factor(absnt$Son)



#------------------------------------Missing Values Analysis---------------------------------------------------#
#Creating dataframe with missing values present in each variable
missing_val = data.frame(apply(absnt,2,function(x){sum(is.na(x))}))

#Calculating percentage missing value
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(absnt)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]

#selecting missing value imputation method
# Actual Value = 23
# Mean = 26.68
# Median = 25
# KNN = 23


#Mean Method
# absnt$Body.mass.index[is.na(absnt$Body.mass.index)] = mean(absnt$Body.mass.index, na.rm = T)

#Median Method
# absnt$Body.mass.index[is.na(absnt$Body.mass.index)] = median(absnt$Body.mass.index, na.rm = T)

# kNN Imputation
absnt = knnImputation(absnt, k = 3)


#-------------------------------------Outlier Analysis-------------------------------------#
# BoxPlots - Distribution and Outlier Check

# Boxplot for continuous variables
for (i in 1:length(continuous_vars))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (continuous_vars[i]), x = "Absenteeism.time.in.hours"), data = subset(absnt))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=continuous_vars[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of absenteeism for",continuous_vars[i])))
}

# ## Plotting plots together
gridExtra::grid.arrange(gn1,gn2,ncol=2)
gridExtra::grid.arrange(gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,ncol=2)
gridExtra::grid.arrange(gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)


# #Remove outliers using boxplot method


#Replace all outliers with NA and impute
for(i in continuous_vars)
{
  val = absnt[,i][absnt[,i] %in% boxplot.stats(absnt[,i])$out]

  absnt[,i][absnt[,i] %in% val] = NA
}

# Imputing missing values
absnt = knnImputation(absnt,k=3)


#-----------------------------------Feature Selection------------------------------------------#

## Correlation Plot 
corrgram(absnt[,continuous_vars], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")


## Dimension Reduction
absnt = subset(absnt, select = -c(Weight))


####################Quick visualisations###################
ggplot(absnt, aes_string(x = absnt$Education,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "green") + theme_bw() +  xlab("Education") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Reason.for.absence,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "DarkSlateBlue") + theme_bw() +  xlab("reason of absense") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Reason.for.absence,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "DarkSlateBlue") + theme_bw() +  xlab("reason of absense") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$ID,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "Orange") + theme_bw() +  xlab("ID") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Social.smoker,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "Red") + theme_bw() +  xlab("smoker") + ylab('absenteeism')
####################################################################
####################Feature Scaling#####################
#Normality check
hist(absnt$Absenteeism.time.in.hours)

# Updating the continuous and catagorical variable
continuous_vars = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Height', 
                    'Body.mass.index')

catagorical_vars = c('ID','Reason.for.absence','Disciplinary.failure', 
                     'Social.drinker', 'Son', 'Pet', 'Month.of.absence', 'Day.of.the.week', 'Seasons',
                     'Education', 'Social.smoker')


# Normalization
 for(i in continuous_vars)
{
  print(i)
  absnt[,i] = (absnt[,i] - min(absnt[,i]))/(max(absnt[,i])-min(absnt[,i]))
}

# Creating dummy variables for categorical variables

absnt = dummy.data.frame(absnt, catagorical_vars)


#------------------------------------------Model Development--------------------------------------------#
#Cleaning the environment
rmExcept("absnt")

#Divide data into train and test using stratified sampling method
train.index = sample(1:nrow(absnt), 0.8 * nrow(absnt))
train = absnt[ train.index,]
test  = absnt[-train.index,]



#------------------------------------------Linear Regression-------------------------------------------#

#Develop Model on training data
fit_LR = lm(Absenteeism.time.in.hours ~ ., data = train)

#summary
summary(fit_LR)

# prediction for training data
pred_LR_train = predict(fit_LR,train[,-116])

#Lets predict for testing data
pred_LR_test = predict(fit_LR,test[,-116])

#for training data
regr.eval(train[,116],pred_LR_train,stats=c('rmse'))

# For testing data 
regr.eval(test[,116],pred_LR_test,stats=c('rmse'))

#-----------------------------------------Random Forest----------------------------------------------#


#Develop Model on training data
fit_RF = randomForest(Absenteeism.time.in.hours~., data = train)

#Lets predict for training data
pred_RF_train = predict(fit_RF, train[,names(test) != "Absenteeism.time.in.hours"])

#Lets predict for testing data
pred_RF_test = predict(fit_RF,test[,names(test) != "Absenteeism.time.in.hours"])

# For training data
regr.eval(train[,116],pred_RF_train,stats=c('rmse'))

# For testing data 
regr.eval(test[,116],pred_RF_test,stats=c('rmse'))



#----------------------Dimensionality Reduction using PCA-------------------------------#


#principal component analysis
prin_comp = prcomp(train)

#compute standard deviation of each principal component
std_dev = prin_comp$sdev

#compute variance
pr_var = std_dev^2

#proportion of variance explained
prop_varex = pr_var/sum(pr_var)

#cumulative scree plot
plot(cumsum(prop_varex), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "b")

#add a training set with principal components
train.data = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin_comp$x)

# From the above plot selecting 45 components since it explains almost 95+ % data variance
train.data =train.data[,1:45]

#transform test into PCA
test.data = predict(prin_comp, newdata = test)
test.data = as.data.frame(test.data)

#select the first 45 components
test.data=test.data[,1:45]


#------------------------------------Model Development after Dimensionality Reduction--------------------------------------------#

#########################Linear Regression#############################

#Develop Model on training data
fit_LR = lm(Absenteeism.time.in.hours ~ ., data = train.data)



#Lets predict for training data
pred_LR_train = predict(fit_LR, train.data)

#Lets predict for testing data
pred_LR_test = predict(fit_LR,test.data)


# For training data LR
regr.eval(train[,116],pred_LR_train,stats=c('rmse'))

# For testing data LR
regr.eval(test[,116],pred_LR_test,stats=c('rmse'))

plot(cumsum(pred_LR_test), xlab = "Absenteeism in 2011",
     ylab = "trend of Absenteeism Explained",
     type = "b")


###################Random Forest#########################

#Develop Model on training data
fit_RF = randomForest(Absenteeism.time.in.hours~., data = train.data)

#Lets predict for training data
pred_RF_train = predict(fit_RF, train.data)

#Lets predict for testing data
pred_RF_test = predict(fit_RF,test.data)

# For training data RF
regr.eval(train[,116],pred_RF_train,stats=c('rmse'))

# For testing data RF
regr.eval(test[,116],pred_RF_test,stats=c('rmse'))



################out time validation##############

#on compairing performance major of LR model 

# For training data LR
regr.eval(train[,116],pred_LR_train,stats=c('rmse'))

# For testing data LR
regr.eval(test[,116],pred_LR_test,stats=c('rmse'))

#on compairing performance major of LR model 

# For training data RF
regr.eval(train[,116],pred_RF_train,stats=c('rmse'))

# For testing data RF
regr.eval(test[,116],pred_RF_test,stats=c('rmse'))
#######################################################



####################Quick visualisations###################
ggplot(absnt, aes_string(x = absnt$Education,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "green") + theme_bw() +  xlab("Education") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Reason.for.absence,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "DarkSlateBlue") + theme_bw() +  xlab("reason of absense") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Reason.for.absence,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "DarkSlateBlue") + theme_bw() +  xlab("reason of absense") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$ID,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "Orange") + theme_bw() +  xlab("ID") + ylab('absenteeism')


ggplot(absnt, aes_string(x = absnt$Social.smoker,y = absnt$Absenteeism.time.in.hours)) +
  geom_bar(stat="identity",fill =  "Red") + theme_bw() +  xlab("smoker") + ylab('absenteeism')

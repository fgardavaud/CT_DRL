######################################################################
##                                                                  ##                                  
##          DRL completion from DoseWatch database export           ##
##                                                                  ##                                 
######################################################################

# created by François Gardavaud, MPE, M.Sc. Medical imaging department - Tenon University Hospital
# initial creation date : 10/01/2023


###################### set-up section ################################

# Set the projet path to the root level
root.dir = rprojroot::find_rstudio_root_file()

# load readxl package to read easily Excel file with an install condition
if(!require(readxl)){
  install.packages("readxl")
  library(readxl)
}

# load lubridate package to determine patient age from birthdate with an install condition
if(!require(lubridate)){
  install.packages("lubridate")
  library(lubridate)
}

# load doMC package for parallel computing
if(!require(doMC)){
  install.packages("doMC")
  library(doMC)
}

# load doparallel package to determine patient age from birthdate with an install conditon
# if(!require(doParallel)){
#   install.packages("doParallel")
#   library(doParallel)
# }

# load tictoc package to measure running time of R code
if(!require(tictoc)){
  install.packages("tictoc")
  library(tictoc)
}

# load excel package to write results in Excel file
if(!require(openxlsx)){
  install.packages("openxlsx")
  library(openxlsx)
}

# load tidyverse for data science such as data handling and visualization
if(!require(tidyverse)){
  install.packages("tidyverse")
  library(tidyverse)
}

# load grateful to list package used
if(!require(grateful)){
  install.packages("grateful")
  library(grateful)
}

# load prettyR to perform tailored statistical analysis
# if(!require(prettyR)){
#   install.packages("prettyR")
#   library(prettyR)
# }

# # load Rcommander GUI for basic stat analysis
# if(!require(Rcmdr)){
#   install.packages("Rcmdr")
#   library(Rcmdr)
# }

############################### data import section ##################################


## /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\
# unfortunately Excel importation yield to parse errors in date data for instance.
# SO TO AVOID THIS PROBLEM THE USER HAVE TO IMPORT DATA IN CSV FORMAT 
# BY CONVERTING ORIGINAL DATA FROM .XLSX TO .CSV IN EXCEL SOFTWARE
# /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\

# the study is based from patient database extracted between 02/01/2018 to 1/10/2021

tic("to import detailled data in Rstudio")
if(exists("DoseWatch_export")){
  print("raw data importation have already done")
}else{
  
  # operations to correctly read DoseWatch export file *NEW* format only for DoseWatch v3.2.3 and above.
  # if you have DoseWatch v3.1 or under comment the three following lines and uncomment the last previous command line.
  all_content = readLines("data/CT_Tenon_Radiologie_detailed_data_export.csv") # to read the whole file
  skip_content = all_content[-c(1,2,3)] # to suppress the first 3 rows with bad value yield to bad header with read.csv2 function
  DoseWatch_raw_data <- read.csv2(textConnection(skip_content), sep = ";")
  # my_all_data <- read.csv2("data_v2/Interventional_Tenon_Radiologie_detailed_data_export.csv", sep = ";")
  rm(all_content, skip_content) # clean temporary variables
}
toc()




################################## data tailoring section ##################################
# data selection to keep only interested columns for this study
DoseWatch_Selected_data <- DoseWatch_raw_data %>% select(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                                         Model, Patient.birthdate..YYYY.MM.DD.,
                                                         Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                                         Num..Series.Non.Localizer, Local.study.description, Series.description,
                                                         Series.type, Series.protocol.name, KVP..kV., Pitch.factor,
                                                         Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                                         Mean.CTDI.vol..mGy., DLP..mGy.cm., Scanning.length..mm., Total.DLP..mGy.cm., Max.CTDIvol..mGy.) # to select column of interest and keeping the column's name

# convert Study.time, Patient Birthdate, Study date and accession number columns in right format
DoseWatch_Selected_data <- DoseWatch_Selected_data %>%
  mutate(Study.time = hm(Study.time), Patient.birthdate..YYYY.MM.DD. = ymd(Patient.birthdate..YYYY.MM.DD.),
         Study.date..YYYY.MM.DD. = as.POSIXct(Study.date..YYYY.MM.DD.),
         Accession.number = as.numeric(Accession.number),
         Scanning.length.cm = Scanning.length..mm. / 10)

# sort each line by Accession number and then by acquisition hour
DoseWatch_Selected_data <- arrange(DoseWatch_Selected_data, Accession.number, Study.time)

######################## age patient and study date computation #################################
#  instance null vector with appropriate dimension to bind with Study_data
Patient.Age <- rep(0, nrow(DoseWatch_Selected_data))

# Loop with parallelization to calculate patient age in years 
# and add this information to Study_data dataframe
# also have a condition to test global environment object for debugging
tic("for loop with parallelization")
if(exists("Study_data_selected_age")){
  print("patient age computation have already done")
}else{
  cores <- detectCores()
  registerDoMC(cores - 1)
  Patient.Age <- foreach(i = 1:nrow(DoseWatch_Selected_data)) %dopar% {
    #naiss = ymd_hms(DoseWatch_Selected_data[i,4]) # deprecated line as mutate function can convert easily "time" column
    #evt = as.POSIXct(DoseWatch_Selected_data[i,1]) # deprecated line as mutate function can convert easily "time" column
    # by suppressing those 2 previous lines and use mutate function instead => Computing acceleration by a factor 6 !!!!!!
    age = as.period(interval(DoseWatch_Selected_data[i,5], DoseWatch_Selected_data[i,2]))@year
    Patient.Age <- age
  }
}
toc()

Patient.Age <- as.character(Patient.Age)
Study_data_age <-cbind(DoseWatch_Selected_data,Patient.Age)

Study_data_selected_age <- Study_data_age %>% select(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                                     Model, Patient.Age,
                                                     Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                                     Num..Series.Non.Localizer, Local.study.description, Series.description,
                                                     Series.type, Series.protocol.name, KVP..kV., Pitch.factor,
                                                     Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                                     Mean.CTDI.vol..mGy., DLP..mGy.cm., Scanning.length.cm, Total.DLP..mGy.cm., Max.CTDIvol..mGy.)
# clean temporary variables
rm(Patient.Age, cores)

# Extract study year to label graphics titles. The corresponding first line of Study_data_selected_age is retained
Study_year <- year(as.Date(Study_data_selected_age$Study.date..YYYY.MM.DD.[1]))
# create a subfolder (if necessary) to stock the results by year
dir.create(file.path(paste0(root.dir,"/output/",Study_year)))

############### column format conversion #################
# convert patient years in numeric in R basic language
Study_data_selected_age$Patient.Age <- as.numeric(Study_data_selected_age$Patient.Age)

# convert dedicated columns as factor with tidyverse library
Study_data_selected_age <- Study_data_selected_age %>%
  mutate(Series.type = as.factor(Series.type),
         Patient.sex = as.factor(Patient.sex),
         Series.description = as.factor(Series.description),
         Series.protocol.name = as.factor(Series.protocol.name),
         Local.study.description = as.factor(Local.study.description))

############### Data tayloring ############
# select only right BMI for DRL analysis between 18 and 35 in France
Study_data_selected_age <- Study_data_selected_age %>% filter(between(BMI, 18.01, 34.99))

# select only CT helicoidal sequence
Study_data_selected_age <- Study_data_selected_age %>% filter(Series.type == "Spiral")

# remove row without value in series.protocol.name and Series.description columns
Exam_data <- Study_data_selected_age %>% mutate(Series.protocol.name = replace(Series.protocol.name, Series.protocol.name == "", NA))
Exam_data <- Exam_data %>% mutate(Series.description = replace(Series.description, Series.description == "", NA))

# Exam_data %>%
#   # recode empty strings "" by NAs
#   na_if("") %>%
#   # remove NAs
#   na.omit

# remove rows with na value in interested columns
Exam_data <- Exam_data %>% drop_na(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                   Model, Patient.Age,
                                   Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                   Num..Series.Non.Localizer, Local.study.description, Series.description,
                                   Series.type, Series.protocol.name, KVP..kV., Pitch.factor,
                                   Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                   Mean.CTDI.vol..mGy., DLP..mGy.cm., Scanning.length.cm, Total.DLP..mGy.cm., Max.CTDIvol..mGy.)

# remove rows with unknown or other in sex type value
Exam_data <- Exam_data %>% filter(Patient.sex == "MALE" | Patient.sex == "FEMALE")



# keep only protocols with 30 exams or more by scanner (condition in French law)
Exam_data <- Exam_data %>%
  group_by(Model, Series.protocol.name) %>%
  filter(n() > 29)

# scanner list to get the model names
list_scanner <- unique(Exam_data$Model)
# instance counter
j <- 1
# list vectors instance to get each dataframe specific to each CT units with or without acquisition details
list_CT_model_data_dataframe <- vector(mode = "list", length = length(list_scanner))
list_CT_model_data_dataframe_wo_duplicates <- vector(mode = "list", length = length(list_scanner))
# read data by scanner
for (CT_model in list_scanner) {
  # data reading of current CT unit
  CT_buffer <- Exam_data %>% filter(Model == CT_model) # to select column of interest and keeping the column's name
  # list protocol
  list_protocol_buffer <- unique(CT_buffer$Series.protocol.name)
  list_protocol_buffer[list_protocol_buffer==""]<-NA # replace empty value by NA
  list_protocol_buffer <- list_protocol_buffer[!is.na(list_protocol_buffer)] # drop Na value in factor
  # print in terminal the Standard Study Description to consider 
  # print(paste0("In this study you have to consider the following CT protocols for ", CT_model, " to perform DRL analysis (30 exams or more) :"))
  # print(cat(unique(CT_buffer$Series.protocol.name)))
  print(paste0("The CT protocol number for ", CT_model, " used in ", Study_year ," (before IRSN selection rules) is : ", length(unique(CT_buffer$Series.protocol.name))))
  # create dataframe specific to each scanner
  assign(paste0("Exam_data_",CT_model),CT_buffer)
  # create factor list of protocol specific to each scanner
  assign(paste0("CT_Protocols_list_for_",CT_model),list_protocol_buffer)
  # create a list of dataframe by grouping all the data
  list_CT_model_data_dataframe[j] <- list(CT_buffer) # on affecte le dataframe du scanner à la liste des dataframe
  # create a list of dataframe by grouping all the data without duplicate Accesion number (keep just one helical sequence by exam)
  list_CT_model_data_dataframe_wo_duplicates[j] <- list(CT_buffer[!duplicated(CT_buffer$Accession.number), ]) # to keep only one row for each exam time.
  j <- j +1
}
rm(CT_buffer, list_protocol_buffer, CT_model, j) # on enlève les variables tampons de la boucle for


###### DRL IRSN format file creation ###########

names(list_CT_model_data_dataframe) <- list_scanner # on affecte le bon nom du scanner à la liste des dataframe

# # loop to compute data on each CT model
j <- 1 # instance a counter to select the right name in the list of dataframe and the CT name vector
for(CT_dataframe in list_CT_model_data_dataframe) {
  
  scanner_name <- names(list_CT_model_data_dataframe[j]) # select the current name of the dataframe in the list_CT_model_data_dataframe list
  scanner_name <- gsub("[ ]", "_", scanner_name, perl=TRUE) # replace " " by "_"
  # CT unit subfolder creation if necessary
  if (!dir.exists(paste0(root.dir,"/output/",Study_year,"/",scanner_name))){
    dir.create(paste0(root.dir,"/output/",Study_year,"/",scanner_name))
  }
  
  # To retrieve protocols list of the current CT
  list_Protocols <- unique(CT_dataframe$Series.protocol.name)
  list_Protocols[list_Protocols==""]<-NA # replace empty value by NA
  list_Protocols <- list_Protocols[!is.na(list_Protocols)] # drop Na value in factor
  
  # loop to create subset to contain only data associated for only one exam description in IRSN format
  for (protocol_name in list_Protocols) {
    DRL_data <-  CT_dataframe %>% filter(Series.protocol.name == protocol_name) # to select only current exam description data
    
    
    # Several tests conditions performed to include only correct acquisition regarding the protocol acquisition
    # for head acquisition
    if(str_sub(DRL_data$Series.protocol.name[1], 1, 1) == 1) { # test on the first character of the protocol name to detect the anatomical region
      DRL_data <- DRL_data %>% filter(str_detect(Series.description, 'CRANE')) # get only series with the right anatomical description
      DRL_data <- DRL_data %>% filter(between(Scanning.length.cm, 15, 20)) # get only the right scan lenght based on IRSN and SFR guidelines
    }
    # for sinus acquisition
    else if(str_sub(DRL_data$Series.protocol.name[1], 1, 1) == 2) { # test on the first character of the protocol name to detect the anatomical region
      DRL_data <- DRL_data %>% filter(str_detect(Series.description, 'SINUS')) # get only series with the right anatomical description
      #DRL_data <- DRL_data %>% filter(between(Scanning.length.cm, 15, 20)) # there is a lack of recommandation for scanning lenght
    }
    # for lumbar spine acquisition
    else if(str_sub(DRL_data$Series.protocol.name[1], 1, 1) == 7) { # test on the first character of the protocol name to detect the anatomical region
      DRL_data <- DRL_data %>% filter(str_detect(Series.description, '"RACHIS LOMBAIRE"')) # get only series with the right anatomical description
      DRL_data <- DRL_data %>% filter(between(Scanning.length.cm, 25, 32)) # get only the right scan lenght based on IRSN and SFR guidelines
    }
    
    # for thorax/TAP acquisition
    else if(str_sub(DRL_data$Series.protocol.name[1], 1, 1) == 5) {
      DRL_data <- DRL_data %>% filter(str_detect(Series.protocol.name, 'THORAX|HEMOPTYSIE|EMBOLIE|TAP|ABDOPELVIS|AP SANS IV|AP-|AP PORTAL|AP ART'))
      DRL_data_thorax <- DRL_data %>% filter(str_detect(Series.description, 'MEDIASTIN|THORAX|EP|PARENCHYME|HEMOPTYSIE|Reperage|CONTROLE|controle') & Scanning.length.cm > 35 & Scanning.length.cm < 45)
      DRL_data_TAP <- DRL_data %>% filter(str_detect(Series.description, '^TAP') & Scanning.length.cm > 60 & Scanning.length.cm < 75)
      DRL_data_AP <- DRL_data %>% filter(str_detect(Series.description, '^AP|ABDOPELV|ABDOPELVIS') & Scanning.length.cm > 45 & Scanning.length.cm < 55)
    }  
    
    #  for AP acquisition
    else if(str_sub(DRL_data$Series.protocol.name[1], 1, 1) == 6) {
      DRL_data <- DRL_data %>% filter(str_detect(Series.protocol.name, 'ABDO|ABDOPELV|ABDOPELVIS|AP|TAP'))
      DRL_data_thorax <- DRL_data %>% filter(str_detect(Series.description, 'MEDIASTIN|THORAX|EP|PARENCHYME') & Scanning.length.cm > 35 & Scanning.length.cm < 45)
      DRL_data_AP <- DRL_data %>% filter(str_detect(Series.description, '^AP|ABDOPELVIS|ABDOPELV|ABDO PELVIS|^ap|Reperage|CONTROLE|controle') & Scanning.length.cm > 45 & Scanning.length.cm < 55)
      DRL_data_TAP <- DRL_data %>% filter(str_detect(Series.description, '^TAP') & Scanning.length.cm > 60 & Scanning.length.cm < 75)
    } 
    else {
      rm(DRL_data)
    }
    
    if(exists("DRL_data") && nrow(DRL_data) > 29 && str_sub(DRL_data$Series.protocol.name[1], 1, 1) != 5  # condition test to keep only protocols with at least 30 acquisitions
       && str_sub(DRL_data$Series.protocol.name[1], 1, 1) != 6){ # and which are not 5.X or 6.X protocols
      DRL_data <- DRL_data %>% 
        # to create column names in regards with IRSN file format (french law)
        mutate(
          messAge = Patient.Age,
          messPoids = Patient.weight..kg.,
          messTaille = Patient.size..cm.,
          messKv = KVP..kV.,
          messPitch = Pitch.factor,
          messAvanceTable = Table.feed.per.rotation..mm.rot.,
          messCollimationNb = Nominal.total.collimation.width..mm.,
          messCollimationVal = Nominal.single.collimation.width..mm.,
          messIdsvCtdi = Mean.CTDI.vol..mGy.,
          messPdl = DLP..mGy.cm.
        )
      # to remove original columns
      DRL_data <- DRL_data %>%  select(-c(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                          Patient.Age,
                                          Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                          Num..Series.Non.Localizer, Local.study.description,
                                          Series.type, KVP..kV., Pitch.factor,
                                          Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                          Mean.CTDI.vol..mGy., DLP..mGy.cm., Total.DLP..mGy.cm., Max.CTDIvol..mGy.))
      protocol_name <- gsub("[/. ]", "_", protocol_name, perl=TRUE) # replace /" and "." by "_" 
      protocol_name <- gsub("[()+]", "", protocol_name, perl=TRUE) # replace "(", ")" and "+" by "" 
      #scanner_name <- gsub("[ ]", "_", scanner_name, perl=TRUE) # replace " " by "_"
      DRL_name <- paste0("DRL_data_",protocol_name,"_for_", scanner_name, sep = "") # to generate current exam description name for data
      assign(DRL_name, DRL_data) # to assign data from DRL_data to DRL_name
      path_name <- paste("output/",Study_year, "/", scanner_name, "/", DRL_name, "_", Study_year, ".csv", sep ="") # create path name to save Excel file in csv format
      write.csv2(subset(DRL_data, select=-c(Model,Series.description,Series.protocol.name,Scanning.length.cm)),
                 file=path_name, row.names = FALSE, fileEncoding = "UTF-8") # save .csv DRL file with IRSN format (french law)
    }
    
    # part to deal with DRL dataframe specific to thorax acquisition
    if(exists("DRL_data_thorax") && nrow(DRL_data_thorax) > 29){
      DRL_data_thorax <- DRL_data_thorax %>% 
        # to create column names in regards with IRSN file format (french law)
        mutate(
          messAge = Patient.Age,
          messPoids = Patient.weight..kg.,
          messTaille = Patient.size..cm.,
          messKv = KVP..kV.,
          messPitch = Pitch.factor,
          messAvanceTable = Table.feed.per.rotation..mm.rot.,
          messCollimationNb = Nominal.total.collimation.width..mm.,
          messCollimationVal = Nominal.single.collimation.width..mm.,
          messIdsvCtdi = Mean.CTDI.vol..mGy.,
          messPdl = DLP..mGy.cm.
        )
      # to remove original columns
      DRL_data_thorax <- DRL_data_thorax %>%  select(-c(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                                        Patient.Age,
                                                        Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                                        Num..Series.Non.Localizer, Local.study.description,
                                                        Series.type, KVP..kV., Pitch.factor,
                                                        Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                                        Mean.CTDI.vol..mGy., DLP..mGy.cm., Total.DLP..mGy.cm., Max.CTDIvol..mGy.))
      protocol_name <- gsub("[/. ]", "_", protocol_name, perl=TRUE) # replace /" and "." by "_" 
      protocol_name <- gsub("[()+]", "", protocol_name, perl=TRUE) # replace "(", ")" and "+" by "" 
      #scanner_name <- gsub("[ ]", "_", scanner_name, perl=TRUE) # replace " " by "_"
      DRL_name <- paste0("DRL_data_for_thorax_",protocol_name,"_for_", scanner_name, sep = "") # to generate current exam description name for data
      assign(DRL_name, DRL_data_thorax) # to assign data from DRL_data_thorax to DRL_name
      path_name <- paste("output/",Study_year, "/", scanner_name, "/", DRL_name, "_", Study_year, ".csv", sep ="") # create path name to save Excel file in csv format
      write.csv2(subset(DRL_data_thorax, select=-c(Model,Series.description,Series.protocol.name,Scanning.length.cm)),
                 file=path_name, row.names = FALSE, fileEncoding = "UTF-8") # save .csv DRL file with IRSN format (french law)
    }
    
    # part to deal with DRL dataframe specific to TAP acquisition
    if(exists("DRL_data_TAP") && nrow(DRL_data_TAP) > 29){
      DRL_data_TAP <- DRL_data_TAP %>% 
        # to create column names in regards with IRSN file format (french law)
        mutate(
          messAge = Patient.Age,
          messPoids = Patient.weight..kg.,
          messTaille = Patient.size..cm.,
          messKv = KVP..kV.,
          messPitch = Pitch.factor,
          messAvanceTable = Table.feed.per.rotation..mm.rot.,
          messCollimationNb = Nominal.total.collimation.width..mm.,
          messCollimationVal = Nominal.single.collimation.width..mm.,
          messIdsvCtdi = Mean.CTDI.vol..mGy.,
          messPdl = DLP..mGy.cm.
        )
      # to remove original columns
      DRL_data_TAP <- DRL_data_TAP %>%  select(-c(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                                  Patient.Age,
                                                  Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                                  Num..Series.Non.Localizer, Local.study.description,
                                                  Series.type, KVP..kV., Pitch.factor,
                                                  Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                                  Mean.CTDI.vol..mGy., DLP..mGy.cm., Total.DLP..mGy.cm., Max.CTDIvol..mGy.))
      protocol_name <- gsub("[/. ]", "_", protocol_name, perl=TRUE) # replace /" and "." by "_" 
      protocol_name <- gsub("[()+]", "", protocol_name, perl=TRUE) # replace "(", ")" and "+" by "" 
      #scanner_name <- gsub("[ ]", "_", scanner_name, perl=TRUE) # replace " " by "_"
      DRL_name <- paste0("DRL_data_for_TAP_",protocol_name,"_for_", scanner_name, sep = "") # to generate current exam description name for data
      assign(DRL_name, DRL_data_TAP) # to assign data from DRL_data_TAP to DRL_name
      path_name <- paste("output/",Study_year, "/", scanner_name, "/", DRL_name, "_", Study_year, ".csv", sep ="") # create path name to save Excel file in csv format
      write.csv2(subset(DRL_data_TAP, select=-c(Model,Series.description,Series.protocol.name,Scanning.length.cm)),
                 file=path_name, row.names = FALSE, fileEncoding = "UTF-8") # save .csv DRL file with IRSN format (french law)
    }
    
    # part to deal with DRL dataframe specific to AP acquisition in thoracic protocol
    if(exists("DRL_data_AP") && nrow(DRL_data_AP) > 29){
      DRL_data_AP <- DRL_data_AP %>% 
        # to create column names in regards with IRSN file format (french law)
        mutate(
          messAge = Patient.Age,
          messPoids = Patient.weight..kg.,
          messTaille = Patient.size..cm.,
          messKv = KVP..kV.,
          messPitch = Pitch.factor,
          messAvanceTable = Table.feed.per.rotation..mm.rot.,
          messCollimationNb = Nominal.total.collimation.width..mm.,
          messCollimationVal = Nominal.single.collimation.width..mm.,
          messIdsvCtdi = Mean.CTDI.vol..mGy.,
          messPdl = DLP..mGy.cm.
        )
      # to remove original columns
      DRL_data_AP <- DRL_data_AP %>%  select(-c(Accession.number, Study.date..YYYY.MM.DD., Study.time,
                                                Patient.Age,
                                                Patient.sex, Patient.weight..kg., Patient.size..cm., BMI,
                                                Num..Series.Non.Localizer, Local.study.description,
                                                Series.type, KVP..kV., Pitch.factor,
                                                Table.feed.per.rotation..mm.rot., Nominal.total.collimation.width..mm., Nominal.single.collimation.width..mm.,
                                                Mean.CTDI.vol..mGy., DLP..mGy.cm., Total.DLP..mGy.cm., Max.CTDIvol..mGy.))
      protocol_name <- gsub("[/. ]", "_", protocol_name, perl=TRUE) # replace /" and "." by "_" 
      protocol_name <- gsub("[()+]", "", protocol_name, perl=TRUE) # replace "(", ")" and "+" by "" 
      #scanner_name <- gsub("[ ]", "_", scanner_name, perl=TRUE) # replace " " by "_"
      DRL_name <- paste0("DRL_data_for_AP_",protocol_name,"_for_", scanner_name, sep = "") # to generate current exam description name for data
      assign(DRL_name, DRL_data_AP) # to assign data from DRL_data_AP to DRL_name
      path_name <- paste("output/",Study_year, "/", scanner_name, "/", DRL_name, "_", Study_year, ".csv", sep ="") # create path name to save Excel file in csv format
      write.csv2(subset(DRL_data_AP, select=-c(Model,Series.description,Series.protocol.name,Scanning.length.cm)),
                 file=path_name, row.names = FALSE, fileEncoding = "UTF-8") # save .csv DRL file with IRSN format (french law)
    }
    
    # /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\
    # /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\
    
    # You have to open in Excel the output .csv files and convert them in MS-DOS format even if it seems to be the case
    
    # /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\
    # /!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\/!\
    
    # clean Global environment 
    if(exists("DRL_data")){
      rm(DRL_data)
    }
    if(exists("DRL_name")){
      rm(DRL_name)
    }
    if(exists("path_name")){
      rm(path_name)
    }
    if(exists("DRL_data_thorax")){
      rm(DRL_data_thorax)
    }
    if(exists("DRL_data_TAP")){
      rm(DRL_data_TAP)
    }
    if(exists("DRL_data_AP")){
      rm(DRL_data_AP)
    }
  }
  j <- j+1 # to loop on CT units
}

rm(j)

print("//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\")
print("DRL extraction is finished") 
print("Becareful you have to open in Excel the output .csv files and convert them in MS-DOS format even if it seems to be the case")  
print("//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\//!\\")


############### Graphical analysis #################
names(list_CT_model_data_dataframe_wo_duplicates) <- list_scanner # the correct scanner name is assigned to the dataframe list

# # loop to compute data on each CT model
j <- 1 # instance a counter to select the right name in the list of dataframe and the CT name vector

for (CT_dataframe in list_CT_model_data_dataframe_wo_duplicates) {
  
  scanner_name <- names(list_CT_model_data_dataframe[j]) # select the current name of the dataframe in the list_CT_model_data_dataframe list
  
  # Retrieve the list of scanner protocols to be studied
  list_Description <- unique(CT_dataframe$Local.study.description)
  list_Description[list_Description==""]<-NA # replace empty value by NA
  list_Description[list_Description=="9:DIVERS"]<-NA # replace value by NA
  list_Description[list_Description=="9:RIS"]<-NA # replace value by NA
  list_Description <- list_Description[!is.na(list_Description)] # drop NA value in factor
  
  # We retrieve the list of the 10 most frequent exam descriptions then we filter on the dataframe
  Frequent_descriptions_list <- CT_dataframe %>%
    group_by(Local.study.description) %>% count(Local.study.description) %>% # count the number of occurrences for the descriptions
    arrange(desc(n)) %>% # data is sorted in descending order
    head(10) %>% # select the first ten lines (i.e. the 10 most frequent)
    pull(Local.study.description) # a vector of study descriptions is created
  CT_dataframe_description_filtered <- CT_dataframe %>% filter(Local.study.description %in% Frequent_descriptions_list) # filter according to the list of frequent descriptions
  
  # Plot histogram for each frequent study description
  CT_dataframe_description_filtered %>%
    ggplot() +
    aes(x = Local.study.description) +
    geom_bar(fill = "#D96748") +
    labs(x = "Descriptions d'examen", y = "Nombre d'actes", 
         title = paste0("Nombre d'examens selon les 10 descriptions cliniques les plus fréquentes du scanner ",scanner_name," pour des patients d'IMC standard pour l'année ", Study_year),
         # caption = "histo_exam"
    ) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("Frequent_Local_Study_Description_histogram_for_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  # Plot CTDIvol max boxplot for each frequent exam description
  CT_dataframe_description_filtered %>%
    ggplot() +
    aes(x = Local.study.description, y = Max.CTDIvol..mGy.) +
    geom_boxplot(shape = "circle", fill = "#D96748") +
    labs(x = "Descriptions d'examen", y = "CTDIvol maximal par examen (mGy)", 
         title = paste0("Distribution du CTDIvol max pour 10 examens les plus fréquent du scanner ",scanner_name, " pour l'année ", Study_year)) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("CTDIvolmax_Local_Study_Description_boxplot_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  # Plot total DLP boxplot for each frequent exam description
  CT_dataframe_description_filtered %>%
    ggplot() +
    aes(x = Local.study.description, y = Total.DLP..mGy.cm.) +
    geom_boxplot(shape = "circle", fill = "#D96748") +
    labs(x = "Descriptions d'examen", y = "PDL total par examen (mGy.cm)", 
         title = paste0("Distribution du PDL total pour 10 examens les plus fréquent du scanner ",scanner_name, " pour l'année ", Study_year)) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("Total_DLP_Local_Study_Description_boxplot_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  
  
  
  
  
  # Retrieve the list of the 6 most frequent examination protocols then we filter on the dataframe
  Frequent_protocols_list <- CT_dataframe %>%
    group_by(Series.protocol.name) %>% count(Series.protocol.name) %>% # count the number of occurrences for the descriptions
    arrange(desc(n)) %>% # data is sorted in descending order
    head(6) %>% # select the first 6 lines (i.e. the 6 most frequent)
    pull(Series.protocol.name) # a vector of study descriptions is created
  CT_dataframe_protocol_filtered <- CT_dataframe %>% filter(Series.protocol.name %in% Frequent_protocols_list) # filter according to the list of frequent descriptions
  
  # Plot histogram for each frequent study description
  CT_dataframe_protocol_filtered %>%
    ggplot() +
    aes(x = Series.protocol.name) +
    geom_bar(fill = "#D96748") +
    labs(x = "Protocoles", y = "Nombre d'actes", 
         title = paste0("Nombre d'examens selon les 6 protocoles les plus fréquents du scanner ",scanner_name," pour des patients d'IMC standard pour l'année ", Study_year),
         #caption = "histo_exam"
    ) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("Frequent_Protocols_histogram_for_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  # Plot CTDIvol max boxplot for each frequent exam description
  CT_dataframe_protocol_filtered %>%
    ggplot() +
    aes(x = Series.protocol.name, y = Max.CTDIvol..mGy.) +
    geom_boxplot(shape = "circle", fill = "#D96748") +
    labs(x = "Protocoles", y = "CTDIvol maximal par examen (mGy)", 
         title = paste0("Distribution du CTDIvol max pour les 6 protocoles les plus fréquents du scanner ",scanner_name, " pour l'année ", Study_year)) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("CTDIvolmax_Protocols_boxplot_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  # Plot total DLP boxplot for each frequent exam description
  CT_dataframe_protocol_filtered %>%
    ggplot() +
    aes(x = Series.protocol.name, y = Total.DLP..mGy.cm.) +
    geom_boxplot(shape = "circle", fill = "#D96748") +
    labs(x = "Protocoles", y = "PDL total par examen (mGy.cm)", 
         title = paste0("Distribution du PDL total pour les 6 protocoles les plus fréquents du scanner ",scanner_name, " pour l'année ", Study_year)) +
    theme_gray()
  ggsave(path = paste0("output/",Study_year), filename = paste0("Total_DLP_Protocols_boxplot_",scanner_name,"_in_",Study_year,".png"), width = 16)
  
  j <- j+1 # to loop on CT units
}


############### Statistical analysis #################

# # loop to compute data on each CT model
j <- 1 # instance a counter to loop on CT units

for (CT_dataframe in list_CT_model_data_dataframe_wo_duplicates) {
  
  scanner_name <- names(list_CT_model_data_dataframe[j]) # select the current name of the dataframe in the list_CT_model_data_dataframe list
  
  # Retrieve the list of the 10 most frequent exam descriptions then we filter on the dataframe
  Frequent_descriptions_list <- CT_dataframe %>%
    group_by(Local.study.description) %>% count(Local.study.description) %>% # count the number of occurrences for the descriptions
    arrange(desc(n)) %>% # data is sorted in descending order
    head(10) %>% # select the first ten lines (i.e. the 10 most frequent)
    pull(Local.study.description) # a vector of study descriptions is created
  CT_dataframe_description_filtered <- CT_dataframe %>% filter(Local.study.description %in% Frequent_descriptions_list) # filter according to the list of frequent descriptions
  
  
  # get statistics (mean, sd, max, min) round at unit for 10 more frequent exam description
  Local_DRL_study_description <- CT_dataframe_description_filtered %>%
    group_by(Local.study.description) %>%
    summarise(
      # stats for Total DLP in mGy.cm
      mean_Total_DLP_mGy.cm = round(mean(Total.DLP..mGy.cm., na.rm = TRUE),0),
      med_Total_DLP_mGy.cm = round(median(Total.DLP..mGy.cm., na.rm = TRUE),0),
      sd_Total_DLP_mGy.cm = round(sd(Total.DLP..mGy.cm., na.rm = TRUE),0),
      max_Total_DLP_mGy.cm = round(max(Total.DLP..mGy.cm., na.rm = TRUE),0),
      min_Total_DLP_mGy.cm = round(min(Total.DLP..mGy.cm., na.rm = TRUE),0),
      # stats for max CTDIvol in mGy
      mean_max_CTDIvol_mGy = round(mean(Max.CTDIvol..mGy., na.rm = TRUE),0),
      med_max_CTDIvol_mGy = round(median(Max.CTDIvol..mGy., na.rm = TRUE),0),
      sd_max_CTDIvol_mGy = round(sd(Max.CTDIvol..mGy., na.rm = TRUE),0),
      max_max_CTDIvol_mGy = round(max(Max.CTDIvol..mGy., na.rm = TRUE),0),
      min_max_CTDIvol_mGy = round(min(Max.CTDIvol..mGy., na.rm = TRUE),0),
      # stats for kV
      mean_kV = round(mean(KVP..kV., na.rm = TRUE),0),
      med_kV = round(median(KVP..kV., na.rm = TRUE),0),
      sd_kV = round(sd(KVP..kV., na.rm = TRUE),0),
      max_kV = round(max(KVP..kV., na.rm = TRUE),0),
      min_kV = round(min(KVP..kV., na.rm = TRUE),0),
      # stats for Pitch
      mean_Pitch = round(mean(Pitch.factor, na.rm = TRUE),0),
      med_Pitch = round(median(Pitch.factor, na.rm = TRUE),0),
      sd_Pitch = round(sd(Pitch.factor, na.rm = TRUE),0),
      max_Pitch = round(max(Pitch.factor, na.rm = TRUE),0),
      min_Pitch = round(min(Pitch.factor, na.rm = TRUE),0),
      # stats for acquisition (only spiral, smartprep & smartstep) number data
      mean_Acq_Num = round(mean(Num..Series.Non.Localizer, na.rm = TRUE),0),
      med_Acq_Num  = round(median(Num..Series.Non.Localizer, na.rm = TRUE),0),
      sd_Acq_Num = round(sd(Num..Series.Non.Localizer, na.rm = TRUE),0),
      max_Acq_Num = round(max(Num..Series.Non.Localizer, na.rm = TRUE),0),
      min_Acq_Num = round(min(Num..Series.Non.Localizer, na.rm = TRUE),0),
      # Exam number
      Exam_number = n(),
      # stats for BMI
      mean_BMI = round(mean(BMI, na.rm = TRUE),0),
      med_BMI = round(median(BMI, na.rm = TRUE),0),
      sd_BMI = round(sd(BMI, na.rm = TRUE),0),
      max_BMI = round(max(BMI, na.rm = TRUE),0),
      min_BMI = round(min(BMI, na.rm = TRUE),0),
    )
  
  write.xlsx(Local_DRL_study_description, paste0('output/',Study_year,'/Study_Description_Local_DRL_for_',scanner_name,'_in_',Study_year,'.xlsx'), sheetName = paste0("Local_DRL_",scanner_name),
             colNames = TRUE, rowNames = FALSE, append = FALSE, overwrite = TRUE) #rowNames = FALSE to suppress the first column with index
  
  
  # Retrieve the list of the 6 most frequent examination protocols then we filter on the dataframe
  Frequent_protocols_list <- CT_dataframe %>%
    group_by(Series.protocol.name) %>% count(Series.protocol.name) %>% # count the number of occurrences for the descriptions
    arrange(desc(n)) %>% # data is sorted in descending order
    head(6) %>% # select the first 6 lines (i.e. the 6 most frequent)
    pull(Series.protocol.name) # a vector of study descriptions is created
  CT_dataframe_protocol_filtered <- CT_dataframe %>% filter(Series.protocol.name %in% Frequent_protocols_list) # filter according to the list of frequent descriptions
  
  # get statistics (mean, sd, max, min) round at unit for 6 more frequent protocols
  Local_DRL_protocols <- CT_dataframe_protocol_filtered %>%
    group_by(Series.protocol.name) %>%
    summarise(
      # stats for Total DLP in mGy.cm
      mean_Total_DLP_mGy.cm = round(mean(Total.DLP..mGy.cm., na.rm = TRUE),0),
      med_Total_DLP_mGy.cm = round(median(Total.DLP..mGy.cm., na.rm = TRUE),0),
      sd_Total_DLP_mGy.cm = round(sd(Total.DLP..mGy.cm., na.rm = TRUE),0),
      max_Total_DLP_mGy.cm = round(max(Total.DLP..mGy.cm., na.rm = TRUE),0),
      min_Total_DLP_mGy.cm = round(min(Total.DLP..mGy.cm., na.rm = TRUE),0),
      # stats for max CTDIvol in mGy
      mean_max_CTDIvol_mGy = round(mean(Max.CTDIvol..mGy., na.rm = TRUE),0),
      med_max_CTDIvol_mGy = round(median(Max.CTDIvol..mGy., na.rm = TRUE),0),
      sd_max_CTDIvol_mGy = round(sd(Max.CTDIvol..mGy., na.rm = TRUE),0),
      max_max_CTDIvol_mGy = round(max(Max.CTDIvol..mGy., na.rm = TRUE),0),
      min_max_CTDIvol_mGy = round(min(Max.CTDIvol..mGy., na.rm = TRUE),0),
      # stats for kV
      mean_kV = round(mean(KVP..kV., na.rm = TRUE),0),
      med_kV = round(median(KVP..kV., na.rm = TRUE),0),
      sd_kV = round(sd(KVP..kV., na.rm = TRUE),0),
      max_kV = round(max(KVP..kV., na.rm = TRUE),0),
      min_kV = round(min(KVP..kV., na.rm = TRUE),0),
      # stats for Pitch
      mean_Pitch = round(mean(Pitch.factor, na.rm = TRUE),0),
      med_Pitch = round(median(Pitch.factor, na.rm = TRUE),0),
      sd_Pitch = round(sd(Pitch.factor, na.rm = TRUE),0),
      max_Pitch = round(max(Pitch.factor, na.rm = TRUE),0),
      min_Pitch = round(min(Pitch.factor, na.rm = TRUE),0),
      # stats for acquisition (only spiral, smartprep & smartstep) number data
      mean_Acq_Num = round(mean(Num..Series.Non.Localizer, na.rm = TRUE),0),
      med_Acq_Num  = round(median(Num..Series.Non.Localizer, na.rm = TRUE),0),
      sd_Acq_Num = round(sd(Num..Series.Non.Localizer, na.rm = TRUE),0),
      max_Acq_Num = round(max(Num..Series.Non.Localizer, na.rm = TRUE),0),
      min_Acq_Num = round(min(Num..Series.Non.Localizer, na.rm = TRUE),0),
      # Exam number
      Exam_number = n(),
      # stats for BMI
      mean_BMI = round(mean(BMI, na.rm = TRUE),0),
      med_BMI = round(median(BMI, na.rm = TRUE),0),
      sd_BMI = round(sd(BMI, na.rm = TRUE),0),
      max_BMI = round(max(BMI, na.rm = TRUE),0),
      min_BMI = round(min(BMI, na.rm = TRUE),0),
    )
  
  write.xlsx(Local_DRL_protocols, paste0('output/',Study_year,'/Protocols_Local_DRL_for_',scanner_name,'_in_',Study_year,'.xlsx'), sheetName = paste0("Local_DRL_",scanner_name),
             colNames = TRUE, rowNames = FALSE, append = FALSE, overwrite = TRUE) #rowNames = FALSE to suppress the first column with index
  
  
  rm(Local_DRL_protocols, Local_DRL_study_description)
  j <- j+1 # to loop on CT units
}



## ####################### Miscellaneous #####################################
# Create word document to list package citation
#cite_packages(out.format = "docx", out.dir = file.path(paste0(root.dir,"/output/")))
cite_packages(out.format = "docx")
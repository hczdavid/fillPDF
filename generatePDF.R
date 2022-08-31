

generatePDF <- function(mydata, act_neg, individual, mypath =NULL){
  
  
  print("Please wait! I am working hard on it!")
  # set the work directory 
  path <-getwd()
  #----------------------------#
  # Read the blank form 
  #----------------------------#
  
  # get the path for blank files
  b7001 <- paste0(path, "/Source/XXX")
  b7002 <- paste0(path, "/Source/XXX")
  b7005 <- paste0(path, "/Source/XXX")
  bcl   <- paste0(path, "/Source/XXX")
  check <- paste0(path, "/Source/XXX")
  
  # get the fields for each file
  fields7001 <- staplr::get_fields(b7001, convert_field_names = T)
  fields7002 <- staplr::get_fields(b7002, convert_field_names = T)
  fields7005 <- staplr::get_fields(b7005, convert_field_names = T)
  fields_bcl <- staplr::get_fields(bcl,   convert_field_names = T)
  fields_cl  <- staplr::get_fields(check, convert_field_names = T)
  
  
  
  
  #-----------------------------------------------------#
  # build the dictionary between form and audit tools
  #-----------------------------------------------------#
  
  # get letter number index for Excel
  letter_index        <- c(LETTERS,paste0("A",LETTERS), paste0("B",LETTERS), paste0("C",LETTERS))
  number_index        <- seq_along(letter_index)
  names(number_index) <- letter_index
  
  #get 7000's Time Standard Holiday and finalize(New)
  
  timestandard_cy3  <- xlsx::read.xlsx(file = paste0(path, "/Source/XXX"), sheetIndex = 1, header = F)
  cy3time <- format(timestandard_cy3, "%m/%d/%Y")
  
  
  # active file dictionary 
  fname_b7002    <- names(fields7002)
  active_top_dic <- data.frame(col_ind = number_index[c("E","A","B","D","U","J","K","C","M","P","Q")],
                               fieldname = fname_b7002[c(1,3:7,9,10,16:18)])
  active_err_dic <- data.frame(col_ind = number_index[c("V","X","AA","AC","AE","AG","AI","AK","AM","AO","AQ","AS","AU",
                                                        "AY","BB","BD","BF","BH","BJ","BP","BR","BT","BV","BX")],
                               EE = fname_b7002[c(24:45,73,46)], IC = fname_b7002[c(50:72,74)])
  
  
  
  # negative file dictionary
  negative_top_dic <-  data.frame(col_ind = number_index[c("E","A","B","D","V","J","K","C","M","AA")],
                                  fieldname = fname_b7002[c(1,3:7,9:10,16,19)])
  negative_err_dic <-  data.frame(col_ind = number_index[c("AH","AI","AJ","AM")],IC = fname_b7002[c(77:80)])
  
  #inquiry file dictionary
  inquiry_top_dic <-  data.frame(col_ind = number_index[c("E","A","B","D","G","C","H")],
                                 fieldname = fname_b7002[c(1,3:6,10,16)])
  
  
  if(is.null(mypath)){
    
    ct  <- mydata[1,1]
    ct  <- strsplit(ct, split = " ")[[1]][2]
    
    cmonth <- mydata[1,5]
    cmonth <- strsplit(cmonth, split = "/")[[1]]
    cmonth <- paste(cmonth[1], cmonth[2], sep = "-")
    mypath <- paste0("XXX",ct,"/Final Findings/",cmonth)
    
  }
  
  if(act_neg == "QA ACTIVE REVIEW"){
    
    active_case <- mydata
    
    active_case[number_index["A"]]     <- sapply(strsplit(active_case[,number_index["A"]],split = " "),function(x){x[2]})
    
    if(active_case[1,number_index["A"]] == "New"){
      active_case[number_index["A"]]     <- "New Hanover"
    }
    
    
    active_case[number_index["M"]]     <- format(active_case[number_index["M"]], "%m/%d/%Y")
    
    active_case <- active_case[active_case$QARN %in% individual,]
    active_QARN   <-sapply(strsplit(active_case$QARN, split = " "), function(x){x[1]})
    
  
    
    active_case[,12] <- toupper(active_case[,12])
    active_case[,92] <- toupper(active_case[,92])
    
    
    
    for (kk in 1:nrow(active_case)){ # loop for all active case
      
      
      # copy an empty form
      temp_fields_b7002  <- fields7002
      temp_fields_b7005  <- fields7005
      temp_fields_b7001  <- fields7001
      temp_fields_bcl    <- fields_bcl
      temp_fields_cl     <- fields_cl
      
      #----------------------------#
      # fill Cover letter
      #----------------------------#
      
      
      temp_fields_bcl[["1-Date"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")#DATE
      temp_fields_bcl[["2-Auditor"]][["value"]] <- active_case[number_index["B"]][kk,]#Auditor
      temp_fields_bcl[["3a-App#"]][["value"]]   <- active_case[number_index["K"]][kk,]#3a-App# 
      temp_fields_bcl[["3-App#"]][["value"]]    <- active_case[number_index["J"]][kk,]#3-App#
      temp_fields_bcl[["4-QARN"]][["value"]]   <- active_case[number_index["C"]][kk,]#4-QARN
      
      #----------------------------#
      # fill Check list
      #----------------------------#
      temp_fields_cl[["1 QARN"]][["value"]]       <- active_case[number_index["C"]][kk,]
      temp_fields_cl[["Sample Month"]][["value"]] <- active_case[number_index["E"]][kk,]
      temp_fields_cl[["7001 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
      temp_fields_cl[["7005 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
      
      #-----------------------------------------#
      # fill the top of 7002 (and 7001/7005) file
      #-----------------------------------------#
      
      # 1. two blank not from the original file
      temp_fields_b7002[["Date"]][["value"]]       <- format(Sys.Date(), "%m/%d/%Y")
      temp_fields_b7002[["Prog/Class"]][["value"]] <- paste(active_case$Program, active_case$Class,sep = "/")[kk]
      
      # 2. fill the text box
      for(ll in 1:nrow(active_top_dic)){
        temp_fields_b7002[[active_top_dic[ll,2]]][["value"]] <- active_case[active_top_dic[ll,1]][kk,]
      }
      
      # 3. fill the check box
      # check box for type of action
      if(active_case[kk,number_index["L"]]=="APP - ONGOING" | active_case[kk,number_index["L"]]== "APP - RETRO"){
        temp_fields_b7002[["App/Reapp"]][["value"]] <- "Yes"
      }else if(active_case[kk,number_index["L"]]== "REDETERM"){
        temp_fields_b7002[["Redeterm"]][["value"]]  <- "Yes"
      }
      
      # check box for case finding
      if(active_case[kk,number_index["CN"]] == "CORRECT"){ 
        temp_fields_b7002[["CORRECT"]][["value"]]   <- "Yes"
      }else if(active_case[kk,number_index["CN"]]=="ELIG ERROR"){
        temp_fields_b7002[["EE"]][["value"]]        <- "Yes"
      }else if (active_case[kk,number_index["CN"]]=="ELIG ERROR W/ IC"){
        temp_fields_b7002[["EE w/ IC"]][["value"]]  <- "Yes"
      }else if (active_case[kk,number_index["CN"]]=="IC ONLY"){
        temp_fields_b7002[["IC ONLY"]][["value"]]   <- "Yes"
      }
      
      # Write out the cover letter and 7002 for CORRECT CASE.
      if(active_case[kk,number_index["CN"]] == "CORRECT"){
        
        
        
        filepath <- paste0(mypath,"/",trimws(active_case$QARN[kk]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl, convert_field_names=T)# there is name problem of the file
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)# there is name problem of the file
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
      }else{
        
        #----------------------------------------------#
        # fill the error part of 7002 file
        #----------------------------------------------#
        
        for(rr in 1:nrow(active_err_dic)){
          if(toupper(active_case[kk, active_err_dic[rr,1]]) == "INTERNAL CONTROL"){ 
            temp_fields_b7002[[active_err_dic[rr,3]]][["value"]]   <- "Yes"
          }else if(toupper(active_case[kk, active_err_dic[rr,1]])=="ERROR"){
            temp_fields_b7002[[active_err_dic[rr,2]]][["value"]]   <- "Yes"
          }
        } # added toupper
        
        
        # fill the IC check box for time line, notice and stop processing timer
        if(active_case[kk,number_index["CN"]] == "ELIG ERROR W/ IC" | active_case[kk,number_index["CN"]] == "IC ONLY"){
          if(active_case[kk, number_index["CB"]] == "INTERNAL CONTROL"){ 
            temp_fields_b7002[["IC-Timeliness"]][["value"]]   <- "Yes"}
          if(active_case[kk, number_index["CD"]] == "INTERNAL CONTROL"){ 
            temp_fields_b7002[["IC-Notice" ]][["value"]]   <- "Yes"}
          if(active_case[kk, number_index["CF"]] == "INTERNAL CONTROL"){ 
            temp_fields_b7002[["IC-SPT"]][["value"]]   <- "Yes"}
        }
        
        # fill page 2 of 7002
        temp_fields_b7002[["Page2Narrative"]][["value"]]   <- active_case[kk, number_index["CO"]]
        temp_fields_b7002[["Page2ManualRef"]][["value"]]   <- active_case[kk, number_index["CP"]]
        
        
        #-----------------------#
        # fill the 7001 file
        #-----------------------#
        
        # Text50:  Response due data
        # Text24:  current data
        # Text24a: Auditor name
        # Text23:  QA Rev#
        # Text23a: County name
        
        temp_fields_b7001[["Text50"]][["value"]]    <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
        temp_fields_b7001[["Text24"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7001[["Text24a"]][["value"]]   <- active_case[number_index["B"]][kk,]
        temp_fields_b7001[["Text23"]][["value"]]    <- active_case[number_index["C"]][kk,]
        temp_fields_b7001[["Text23a"]][["value"]]   <- active_case[number_index["A"]][kk,]
        
        #-----------------------#
        # fill the 7005 file
        #-----------------------#
        
        # Text1a: Submit due date
        # Text5:  current data
        # Text1:  QA Rev#
        # Text14: County name
        
        temp_fields_b7005[["Text1a"]][["value"]]   <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
        temp_fields_b7005[["Text5"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7005[["Text1"]][["value"]]    <- active_case[number_index["C"]][kk,]
        temp_fields_b7005[["Text14"]][["value"]]   <- active_case[number_index["A"]][kk,]
        
        #------------------------------#
        # Write out file for error case
        #------------------------------#
        filepath <- paste0(mypath,"/", trimws(active_case$QARN[kk]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_7001 <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_7005 <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", active_QARN[kk], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl,   convert_field_names=T)
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)
        set_fields(b7001, output_filepath = outfilename_7001, temp_fields_b7001, convert_field_names=T)
        set_fields(b7005, output_filepath = outfilename_7005, temp_fields_b7005, convert_field_names=T)
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
        
      } # end error case if statement
    } # end for all active loop
    
    
    
    
  }else if(act_neg == "QA NEGATIVE REVIEW"){  ########### Start Negative ##################
    
    
    negative_case <- mydata
    
    
    # Correct county name
    negative_case[number_index["A"]]   <- sapply(strsplit(negative_case[,number_index["A"]],split = " "),function(x){x[2]})
    if(negative_case[1,number_index["A"]] == "New"){
       negative_case[number_index["A"]]     <- "New Hanover"
    }
    
    negative_case[number_index["M"]]   <- format(negative_case[number_index["M"]], "%m/%d/%Y")
    #negative_case[number_index["AA"]]   <- format(negative_case[number_index["AA"]], "%m/%d/%Y")
    
    
    
    negative_case <- negative_case[negative_case$QARN %in% individual,]
    negative_QARN <- sapply(strsplit(negative_case$QARN, split = " "), function(x){x[1]})
    
    # change to upper case
    negative_case[,12] <- toupper(negative_case[,12])
    negative_case[,39] <- toupper(negative_case[,39])
    negative_case[,38] <- toupper(negative_case[,38])
    
    
    if(class(negative_case[number_index["AA"]][[1]])=="Date"){
      negative_case[number_index["AA"]] <- as.character.Date(negative_case[number_index["AA"]][[1]])
    }
    
    
    for (kk in 1:nrow(negative_case)){ # loop for all negative case
      
      
      
      # if column L is not DENIAL or W/D, leave it as blank
      if(negative_case[number_index["L"]][kk,] %in% c("DENIAL", "W/D")){
        
        aadate <- negative_case[number_index["AA"]][kk,]
        
        if(!grepl("-",aadate)){
          aadate <- as.Date(as.numeric(aadate)-2, origin = "1900-01-01")
        }
        negative_case[number_index["AA"]][kk,]   <- format(as.Date(aadate), "%m/%d/%Y") # don't use comma after kk
      }else{negative_case[number_index["AA"]][kk,] <- ""}
      
      # copy an empty form
      temp_fields_b7002  <- fields7002
      temp_fields_b7005  <- fields7005
      temp_fields_b7001  <- fields7001
      temp_fields_bcl    <- fields_bcl
      temp_fields_cl     <- fields_cl
      
      #----------------------------#
      # fill Cover letter
      #----------------------------#
      
      temp_fields_bcl[["1-Date"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
      temp_fields_bcl[["2-Auditor"]][["value"]] <- negative_case[number_index["B"]][kk,]
      temp_fields_bcl[["3a-App#"]][["value"]]  <- negative_case[number_index["K"]][kk,]#3a-App# 
      temp_fields_bcl[["3-App#"]][["value"]]    <- negative_case[number_index["J"]][kk,]#3-App#
      temp_fields_bcl[["4-QARN"]][["value"]]    <- negative_case[number_index["C"]][kk,]#4-QARN
      
      
      #----------------------------#
      # fill Check list
      #----------------------------#
      temp_fields_cl[["1 QARN"]][["value"]]       <- negative_case[number_index["C"]][kk,]
      temp_fields_cl[["Sample Month"]][["value"]] <- negative_case[number_index["E"]][kk,]
      temp_fields_cl[["7001 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
      temp_fields_cl[["7005 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
      
      #-----------------------------------------#
      # fill the top of 7002 (and 7001/7005) file
      #-----------------------------------------#
      
      # 1. two blank not from the original file
      temp_fields_b7002[["Date"]][["value"]]       <- format(Sys.Date(), "%m/%d/%Y")
      #temp_fields_b7002[["Prog/Class"]][["value"]] <- negative_case$Program[kk]
      
      # if negative_case[kk,12] is TERM - PROG CHNG or TERM - TRUE TERM, when we need both program and class
      if(negative_case[kk,12] == "DENIAL" | negative_case[kk,12] == "W/D"){
        temp_fields_b7002[["Prog/Class"]][["value"]] <- negative_case$Program[kk]
      }else{
        temp_fields_b7002[["Prog/Class"]][["value"]] <- paste(negative_case$Program, negative_case$Class,sep = "/")[kk]
      }
      
      # 2. fill the text box
      for(ll in 1:nrow(negative_top_dic)){
        temp_fields_b7002[[negative_top_dic[ll,2]]][["value"]] <- negative_case[negative_top_dic[ll,1]][kk,]
      }
      
      
      
      # 3. fill the check box
      if(negative_case[kk,12]=="DENIAL"){
        temp_fields_b7002[["Denial"]][["value"]]        <- "Yes"
      }else if(negative_case[kk,12]=="W/D"){
        temp_fields_b7002[["Withdrawal"]][["value"]]    <- "Yes"
      }else if(negative_case[kk,12]=="TERM - TRUE TERM"){
        temp_fields_b7002[["TrueTerm"]][["value"]]      <- "Yes"
      }else if(negative_case[kk,12]=="TERM - PROG CHNG"){
        temp_fields_b7002[["Term-ProgChng"]][["value"]] <- "Yes"
      }
      
      # check box for case finding
      if(negative_case[kk,39] == "CORRECT"){ 
        temp_fields_b7002[["CORRECT"]][["value"]]   <- "Yes"
      }else if(negative_case[kk,39]=="ELIG ERROR"){
        temp_fields_b7002[["EE"]][["value"]]        <- "Yes"
      }else if (negative_case[kk,39]=="ELIG ERROR W/ IC"){
        temp_fields_b7002[["EE w/ IC"]][["value"]]  <- "Yes"
      }else if (negative_case[kk,39]=="IC ONLY"){
        temp_fields_b7002[["IC ONLY"]][["value"]]   <- "Yes"
      }
      
      # Write out the cover letter and 7002 for CORRECT CASE.
      if(negative_case[kk,39] == "CORRECT"){
        
        filepath <- paste0(mypath,"/",trimws(negative_case$QARN[kk]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl, convert_field_names=T)# there is name problem of the file
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)# there is name problem of the file
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
        
        
      }else{
        
        #----------------------------------------------#
        # fill the error part of 7002 file
        #----------------------------------------------#
        
        #Internal Control Error
        
        for(ss in 1:(nrow(negative_err_dic)-1)){
          if(toupper(negative_case[kk, negative_err_dic[ss,1]]) == "INTERNAL CONTROL"){  # added to upper
            temp_fields_b7002[[negative_err_dic[ss,2]]][["value"]]  <- "Yes"
          }
        }
        
        if (negative_case[kk, negative_err_dic[4,1]] == "ELIG ERROR W/ IC"|negative_case[kk, negative_err_dic[4,1]]=="IC ONLY"){ 
          temp_fields_b7002[[negative_err_dic[4,2]]][["value"]]  <- "Yes"
        }
        
        
        
        
        
        #Eligibility Error
        if(negative_case[kk, number_index["AL"]] == "EE-IMPROPER DENIAL"){ 
          temp_fields_b7002[["Neg-EE-Denial"]][["value"]]   <- "Yes"
        } else if(negative_case[kk, number_index["AL"]] == "EE-IMPROPER WD"){ 
          temp_fields_b7002[["Neg-EE-WD" ]][["value"]]     <- "Yes"
        } else if(negative_case[kk, number_index["AL"]] == "EE-IMPROPER TERM"){ 
          temp_fields_b7002[[ "Neg-EE-Term"]][["value"]]   <- "Yes"}
        
        # text in the error part
        temp_fields_b7002[["Neg-ErrorText"]][["value"]]    <- negative_case[kk, number_index["AK"]]
        temp_fields_b7002[["Page2Narrative"]][["value"]]   <- negative_case[kk, number_index["AN"]]
        temp_fields_b7002[["Page2ManualRef"]][["value"]]   <- negative_case[kk, number_index["AO"]]
        
        
        
        #-----------------------#
        # fill the 7001 file
        #-----------------------#
        
        # Text50:  Response due data
        # Text24:  current data
        # Text24a: Auditor name
        # Text23:  QA Rev#
        # Text23a: County name
        
        temp_fields_b7001[["Text50"]][["value"]]    <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
        temp_fields_b7001[["Text24"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7001[["Text24a"]][["value"]]   <- negative_case[number_index["B"]][kk,]
        temp_fields_b7001[["Text23"]][["value"]]    <- negative_case[number_index["C"]][kk,]
        temp_fields_b7001[["Text23a"]][["value"]]   <- negative_case[number_index["A"]][kk,]
        
        #-----------------------#
        # fill the 7005 file
        #-----------------------#
        
        # Text1a: Submit due date
        # Text5:  current data
        # Text1:  QA Rev#
        # Text14: County name
        
        temp_fields_b7005[["Text1a"]][["value"]]   <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
        temp_fields_b7005[["Text5"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7005[["Text1"]][["value"]]    <- negative_case[number_index["C"]][kk,]
        temp_fields_b7005[["Text14"]][["value"]]   <- negative_case[number_index["A"]][kk,]
        
        #------------------------------#
        # Write out file for error case
        #------------------------------#
        filepath <- paste0(mypath,"/",trimws(negative_case$QARN[kk]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_7001 <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_7005 <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", negative_QARN[kk], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl,   convert_field_names=T)
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)
        set_fields(b7001, output_filepath = outfilename_7001, temp_fields_b7001, convert_field_names=T)
        set_fields(b7005, output_filepath = outfilename_7005, temp_fields_b7005, convert_field_names=T)
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
        
      } # end error case if statement
    } # end for all negative loop
    
    
  }
  else if(act_neg == "QA INQUIRY REVIEW"){  ########### Start Inquiry ##################
    
    
    inquiry_case <- mydata
    
    
    # Correct county name
    inquiry_case[number_index["A"]]   <- sapply(strsplit(inquiry_case[,number_index["A"]],split = " "),function(x){x[2]})# County Name
    inquiry_case[number_index["H"]]   <- format(inquiry_case[number_index["H"]], "%m/%d/%Y")# Date of Agency Decision
    # Inquiry_case[number_index["AA"]]   <- format(Inquiry_case[number_index["AA"]], "%m/%d/%Y")#AP
    
    
    
    inquiry_case <- inquiry_case[inquiry_case$QARN %in% individual,]
    inquiry_QARN <-sapply(strsplit(inquiry_case$QARN, split = " "), function(x){x[1]})
    

    
    for (tt in 1:nrow(inquiry_case)){ # loop for all negative case
      
      
      # copy an empty form
      temp_fields_b7002  <- fields7002
      temp_fields_b7005  <- fields7005
      temp_fields_b7001  <- fields7001
      temp_fields_bcl    <- fields_bcl
      temp_fields_cl     <- fields_cl
      
      #----------------------------#
      # fill Cover letter
      #----------------------------#
      
      
      
      temp_fields_bcl[["1-Date"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
      temp_fields_bcl[["2-Auditor"]][["value"]] <- inquiry_case[number_index["B"]][tt,]
      temp_fields_bcl[["4-QARN"]][["value"]]   <- inquiry_case[number_index["C"]][tt,]#4-QARN
      
      
      #----------------------------#
      # fill Check list
      #----------------------------#
      temp_fields_cl[["1 QARN"]][["value"]]       <- inquiry_case[number_index["C"]][tt]
      temp_fields_cl[["Sample Month"]][["value"]] <- inquiry_case[number_index["E"]][tt,]
      temp_fields_cl[["7001 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
      temp_fields_cl[["7005 Due Date"]][["value"]]<- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
      
      #-----------------------------------------#
      # fill the top of 7002 (and 7001/7005) file
      #-----------------------------------------#
      temp_fields_b7002[["Inquiry"]][["value"]]  <- "Yes"
      # 1. two blank not from the original file
      temp_fields_b7002[["Date"]][["value"]]     <- format(Sys.Date(), "%m/%d/%Y")
      
      
      # 2. fill the text box
      for(uu in 1:nrow(inquiry_top_dic)){
        temp_fields_b7002[[inquiry_top_dic[uu,2]]][["value"]] <- inquiry_case[inquiry_top_dic[uu,1]][tt,]
      }
      
      
      # check box for case finding
      if(inquiry_case[tt,number_index["Z"]] == "CORRECT"){ 
        temp_fields_b7002[["CORRECT"]][["value"]]   <- "Yes"
      }else if(inquiry_case[tt,number_index["Z"]]=="ELIG ERROR"){
        temp_fields_b7002[["EE"]][["value"]]        <- "Yes"
      }else if (inquiry_case[tt,number_index["Z"]]=="IC ONLY"){
        temp_fields_b7002[["IC ONLY"]][["value"]]   <- "Yes"
      }
      
      # Write out the cover letter and 7002 for CORRECT CASE.
      if(inquiry_case[tt,number_index["Z"]] == "CORRECT"){
        
        filepath <- paste0(mypath,"/",trimws(inquiry_case$QARN[tt]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl, convert_field_names=T)# there is name problem of the file
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)# there is name problem of the file
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
        
        
      }else{
        
        #----------------------------------------------#
        # fill the error part of 7002 file
        #----------------------------------------------#
    
        
        
        #Eligibility Error
        if(inquiry_case[tt, number_index["Y"]] == "EE-IMPROPER INQUIRY"){ 
          temp_fields_b7002[["Neg-EE-Inquiry"]][["value"]]   <- "Yes"
        }
        # text in the error part
        # temp_fields_b7002[["Neg-ErrorText"]][["value"]]    <- negative_case[kk, number_index["AK"]]
        temp_fields_b7002[["Page2Narrative"]][["value"]]   <- inquiry_case[tt, number_index["AA"]]
        temp_fields_b7002[["Page2ManualRef"]][["value"]]   <- inquiry_case[tt, number_index["AB"]]
        
        
        
        #-----------------------#
        # fill the 7001 file
        #-----------------------#
        
        # Text50:  Response due data
        # Text24:  current data
        # Text24a: Auditor name
        # Text23:  QA Rev#
        # Text23a: County name
        
        temp_fields_b7001[["Text50"]][["value"]]    <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),2]
        temp_fields_b7001[["Text24"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7001[["Text24a"]][["value"]]   <- inquiry_case[number_index["B"]][tt,]
        temp_fields_b7001[["Text23"]][["value"]]    <- inquiry_case[number_index["C"]][tt,]
        temp_fields_b7001[["Text23a"]][["value"]]   <- inquiry_case[number_index["A"]][tt,]
        
        #-----------------------#
        # fill the 7005 file
        #-----------------------#
        
        # Text1a: Submit due date
        # Text5:  current data
        # Text1:  QA Rev#
        # Text14: County name
        
        temp_fields_b7005[["Text1a"]][["value"]]   <- cy3time[which(cy3time[,1]==format(Sys.Date(), "%m/%d/%Y")),3]
        temp_fields_b7005[["Text5"]][["value"]]    <- format(Sys.Date(), "%m/%d/%Y")
        temp_fields_b7005[["Text1"]][["value"]]    <- inquiry_case[number_index["C"]][tt,]
        temp_fields_b7005[["Text14"]][["value"]]   <- inquiry_case[number_index["A"]][tt,]
        
        #------------------------------#
        # Write out file for error case
        #------------------------------#
        filepath <- paste0(mypath,"/",trimws(inquiry_case$QARN[tt]))
        dir.create(filepath)
        
        outfilename_bcl  <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_7002 <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_7001 <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_7005 <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        outfilename_cl   <- paste0(filepath,  "/", inquiry_QARN[tt], " XXX.pdf")
        
        
        set_fields(bcl,   output_filepath = outfilename_bcl,  temp_fields_bcl,   convert_field_names=T)
        set_fields(b7002, output_filepath = outfilename_7002, temp_fields_b7002, convert_field_names=T)
        set_fields(b7001, output_filepath = outfilename_7001, temp_fields_b7001, convert_field_names=T)
        set_fields(b7005, output_filepath = outfilename_7005, temp_fields_b7005, convert_field_names=T)
        set_fields(check, output_filepath = outfilename_cl,   temp_fields_cl,    convert_field_names=T)
        
        
        
      } # end error case if statement
    } # end for all inquiry loop
    
    
  }
  return(paste("All files are generated. Please check your folder!", "\n The output path is ",filepath))
  
}


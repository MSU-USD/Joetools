#' Standard Error
#'
#' @param data The vector or dataset over which the standard error is calculated
#'
#' @return The calculated Standard Error
#' @importFrom stats sd
#' @export
#'
#' @examples
#' x=c(10, 20, 30, 40, 50)
#' se(x)
se=function(data){
  return(sd(data, na.rm=T)/sqrt(sum(!is.na(data))))
}

#' Simplifying Reports
#'
#' @param input dataframe to simplify results from
#'
#' @return Simplified interacton reports
#' @import tidyverse
#' @export
#'
#' @examples
simplifyAppend <- function(input) {
  # x=match("Signif",colnames(input))+1
  # y=as.numeric(ncol(input))
  df=input%>%
    filter(!is.nan(Diff))%>%
    mutate(across(colnames(input)[match("Signif",colnames(input))+1]:colnames(input)[length(colnames(input))],
                  ~cut(.,breaks=c( 0,.001,.01,.05, .1,.999,1), labels=c("***" ,"**","*",".","NS","1" ))))%>%
    select(1:Signif,where(~sum((.%in%c("***" ,"**","*",".")))>0))
}

#' Flexible Workbook Creation Function
#'
#' @param df The dataframe to be saved
#' @param filename Name of the file, including the xlsx extension
#' @param sheetBy Option to create sheets based on the values in a particular column of the dataframe
#' @param keepNames The function automatically renames the column names with \code{\link[stringr:case]{stringr::str_to_sentence()}}. Setting this to TRUE will leave your column names unaltered
#' @param overwrite Sets permissions for overwriting old files. Default is TRUE
#'
#' @return Creates a Excel workbook from the supplied dataframe
#' @import tidyverse
#' @import openxlsx
#' @import lazyeval
#' @export
#'
#' @examples
#' library(datasets)
#' library(tidyverse)
#' library(openxlsx)
#' data(iris)
#' wbsave(iris, "Iris report.xlsx")
#' wbsave(iris, "Iris report - by Species.xlsx", sheetBy="Species", keepNames=TRUE)
wbsave=function(df,filename,sheetBy=NULL,keepNames=TRUE, overwrite=T){
  if(keepNames==FALSE) {
    names(df)=str_to_sentence(names(df))
    if(!is.null(sheetBy)) {sheetBy=str_to_sentence(sheetBy)}
  }
  wb=createWorkbook()
  df=as.data.frame(df)
  
  TABLE_COLNAMES_STYLE=createStyle(fontSize=11, fontColour ="#44546A", borderColour = "#8EA9DB",
                                   borderStyle = "thick", border="bottom", textDecoration = c("BOLD"))
  addWorksheet(wb,sheetName = "All")
  writeData(wb, sheet = 1, df)
  addStyle(wb, sheet = 1, style=TABLE_COLNAMES_STYLE, rows=1, cols=1:length(colnames(df)))
  width_vec_header_all <- nchar(colnames(df))  + 2
  width_vec_all <- apply(df, 2, function(x) max(nchar(as.character(x)) + ifelse(grepl("@",x),3,2), na.rm = TRUE)) 
  max_vec_header_all <- pmax(width_vec_all, width_vec_header_all)
  setColWidths(wb, sheet = 1, cols = 1:length(colnames(df)), widths = max_vec_header_all)
  wbAll=NULL
  if(!is.null(sheetBy)) {l=levels(as.factor(df[[sheetBy]]))}
  if(!is.null(sheetBy)) {for (i in l){
    wbdata=NULL
    wbdata=df%>%
      filter(!!as.symbol(sheetBy)==i)
    addWorksheet(wb,sheetName = i)
    writeData(wb, sheet = i, wbdata)
    addStyle(wb, sheet = i, style=TABLE_COLNAMES_STYLE, rows=1, cols=1:length(colnames(wbdata)))
    width_vec_header_i <- nchar(colnames(wbdata))  + 2
    width_vec_i <- apply(wbdata, 2, function(x) max(nchar(as.character(x)) + ifelse(grepl("@",x),3,2), na.rm = TRUE)) 
    max_vec_header_i <- pmax(width_vec_i, width_vec_header_i)
    setColWidths(wb, sheet = i, cols = 1:length(colnames(wbdata)), widths = max_vec_header_i)
  }}
  saveWorkbook(wb, filename, overwrite=overwrite)
}

#' Creates a Report with t-tests for a Vector of Outcomes
#'
#' @param df Dataframe to use for the report
#' @param Measures A vector of string names for columns to take means for. Order will be used to make the report.
#' @param Factor A binary factor which will be used for comparisons.  Must be 1 (Treatment) or 0 (No_Treatment)
#' @param paired Options for running t-test.  "Yes" forces paired, t-tests, "No" assumes no pairing, and "Try" will try a paired t-test, and follow up with a unpaired t-test if it fails to run.
#'
#' @return
#' @import tidyverse
#' @export
#'
#'
#' @examples

report=function(df, Measures, Factor, paired=c("Try","Yes", "No")){
  output=df%>%
    group_by(.data[[Factor]])%>%
    summarise_at(vars(Measures), mean,na.rm=T)%>%
    mutate(Levels =ifelse(.data[[Factor]]==1, "Treatment", "No_Treatment"))%>%
    select(-.data[[Factor]])%>%
    pivot_longer(cols=Measures, names_to = "Measure")%>%
    pivot_wider(names_from = "Levels", values_from = "value")%>%
    mutate(Diff=Treatment-No_Treatment)%>%
    mutate(Measure=factor(.data$Measure,levels=Measures))%>%
    arrange(Measure)


  reportpvaluecount=0
  p=NULL
  for(i in Measures){
    reportpvaluecount=reportpvaluecount+1
    form=paste0(i, "~",Factor)
    if(paired=="Yes"){p[reportpvaluecount]=as.numeric(try(t.test(formula=as.formula(form), data=df, paired=T)$p.value))}
    else if (paired=="No"){p[reportpvaluecount]=as.numeric(try(t.test(formula=as.formula(form), data=df)$p.value))}
    else {p[reportpvaluecount]=tryCatch(as.numeric(t.test(formula=as.formula(form), data=df, paired=T)$p.value), error=function(err)
      p[reportpvaluecount]=as.numeric(try(t.test(formula=as.formula(form), data=df)$p.value)), finally = NA)}

  }

  output$p_value=p

  output=output%>%
    mutate(Signif=ifelse(p_value<.001, "***",ifelse(p_value<.01, "**",
                                                    ifelse(p_value<.05, "*", ifelse(p_value<.1, ".", "")))))
}


#' Appends Interactions to a \code{\link[JoeTools:report]{report()}} Output Dataframe
#'
#' @param report The report dataframe that is being appended
#' @param df The original dataframe used for the report
#' @param Measures A vector of string names for the dependent variable. Order must be the same as the repor
#' @param Factor A binary factor (as a string) which will be used for comparisons.  Must be 1 (Treatment) or 0 (No_Treatment)
#' @param Interaction A factor column name (as a string) that will be used as an interacting independant variable.
#' @param Simplify Removed uninteresting columns. On by devfault
#'
#'
#' @return
#' @import tidyverse
#' @export
#'
#' @examples

appendInteraction=function(report,df, Measures,Factor,Interaction, Simplify=T){
  p=NULL
  anova=NULL
  aov_factors=as.vector(Interaction)
  for(i in Measures){
    form=paste0(i, "~",Factor,"*",paste(aov_factors, collapse="*"))
    anova[[i]]=NULL
    if(!is.na(tryCatch(summary(aov(as.formula(form),data=df)), error=function(err) NA))){
      
    
    anova[[i]]=summary(aov(as.formula(form),data=df))[[1]]%>%
      rownames_to_column("Interactions")%>%
      filter(grepl(paste0(Factor,":"), Interactions))%>%
      mutate(Interactions=str_remove(Interactions,paste0(Factor,":")))%>%
      mutate(Interactions=str_remove_all(Interactions, " "))%>%
      select(Interactions,pvalue=`Pr(>F)` )%>%
      pivot_wider(names_from = Interactions, values_from = "pvalue")%>%
      mutate(Measure=i)
    } 
  }
  anova=bind_rows(anova)
  if(nrow(anova)>0&(!is.null(anova))){
    output=report%>%
    left_join(anova, by="Measure")
    if(Simplify==T){
      output=simplifyAppend(output)
    }
  } else  {
    warning("Error in ANOVA, No comparison added.")
    output=report
  }
}

#' Plotting more than one plot together
#'
#' @param ... Name of each plot
#' @param plotlist A vector with plot names
#' @param cols How many columns of plots you want in your chart
#' @param layout A matrix specifying the layout. If present, 'cols' is ignored.
#'
#' @return
#' @import grid
#' @export
#'
#' @examples
multiplot <- function(..., plotlist=NULL, cols=1, layout=NULL) {
  library(grid)
  
  # Make a list from the ... arguments and plotlist
  plots <- c(list(...), plotlist)
  
  numPlots = length(plots)
  
  # If layout is NULL, then use 'cols' to determine layout
  if (is.null(layout)) {
    # Make the panel
    # ncol: Number of columns of plots
    # nrow: Number of rows needed, calculated from # of cols
    layout <- matrix(seq(1, cols * ceiling(numPlots/cols)),
                     ncol = cols, nrow = ceiling(numPlots/cols))
  }
  
  if (numPlots==1) {
    print(plots[[1]])
    
  } else {
    # Set up the page
    grid.newpage()
    pushViewport(viewport(layout = grid.layout(nrow(layout), ncol(layout))))
    
    # Make each plot, in the correct location
    for (i in 1:numPlots) {
      # Get the i,j matrix positions of the regions that contain this subplot
      matchidx <- as.data.frame(which(layout == i, arr.ind = TRUE))
      
      print(plots[[i]], vp = viewport(layout.pos.row = matchidx$row,
                                      layout.pos.col = matchidx$col))
    }
  }
}

#' Force Flextable to Page
#'
#' @param ft Flextable object
#' @param pgwidth Customizable page width setting
#'
#' @return
#' @import flextable
#' @export
#'
#' @examples
FitFlextableToPage <- function(ft, pgwidth = 6){
  
  ft_out <- ft %>% autofit()
  
  ft_out <- width(ft_out, width = dim(ft_out)$widths*pgwidth /(flextable_dim(ft_out)$widths))
  return(ft_out)
}

#' Produce the Pairwise Comparisons on factors in a group
#'
#' @param df Dataframe with by subject Factor, Independent and Dependent Variable info
#' @param Ind Column with the independant variable
#' @param Dep Column with the dependant variable
#' @param Factor Column over which to create comparisons
#' @param Logit Allows for logistic regression if Logit=TRUE, off by default
#'
#' @return
#' @import tidyverse
#' @export
#'
#' @examples
pairwise <- function(df, Ind, Dep, Factor, Logit=FALSE){
  test=df%>%
    # filter(if(Year!="All"){Cohort==Year} else {Cohort==Cohort})%>%
    filter(!is.na(!!Dep))%>%
    mutate_("Dependant"=Dep, "Independant"=Ind, "Levels"=Factor)%>%
    group_by(Independant, Levels)%>%
    summarize(Mean=mean(Dependant, na.rm=T),Count=n(), Measure=Dep, Dep_Values=list(Dependant))%>%
    filter(Count>2)%>%
    select(-Count)%>%
    ungroup()%>%
    mutate(Independant=ifelse(Independant==1,"Treatment","No_Treatment"))%>%
    pivot_wider(names_from = Independant, values_from = c("Mean", "Dep_Values"))%>%
    rename(No_Treatment=Mean_No_Treatment, Treatment=Mean_Treatment)%>%
    filter(!is.na(No_Treatment)&!is.na(Treatment))%>%
    mutate(Diff=Treatment-No_Treatment)%>%
    group_by(Levels)
  if(nrow(test)==0){
    pvalue=df%>%
      # filter(if(Year!="All"){Cohort==Year} else {Cohort==Cohort})%>%
      filter(!is.na(!!Dep))%>%
      mutate_("Dependant"=Dep, "Independant"=Ind, "Levels"=Factor)%>%
      group_by(Independant, Levels)%>%
      summarize(Mean=mean(Dependant, na.rm=T), Measure=Dep, Dep_Values=list(Dependant))%>%
      # filter(Count>2)%>%
      # select(-Count)%>%
      ungroup()%>%
      mutate(Independant=ifelse(Independant==1,"Treatment","No_Treatment"))%>%
      pivot_wider(names_from = Independant, values_from = c("Mean", "Dep_Values"))%>%
      rename(No_Treatment=Mean_No_Treatment, Treatment=Mean_Treatment)%>%
      filter(!is.na(No_Treatment)&!is.na(Treatment))%>%
      mutate(Diff=Treatment-No_Treatment)%>%
      group_by(Levels)%>%
      mutate(p_value=NA)
  }else if(Logit==TRUE){
    pvalue=test%>%
      mutate(data=map2(Dep_Values_No_Treatment, Dep_Values_Treatment, 
                       ~bind_rows(tibble(Independant=rep(1,length(.y)), Dependant=.y),tibble(Independant=rep(0,length(.x)), Dependant=.x))))%>%
      mutate(p_value=tryCatch({map(data,~summary(glm(Dependant~Independant,family="binomial", data=.x))$coefficients[2,4])},
                              error=function(e) {NA}))%>%
      select(-data)%>%
      mutate(p_value=as.numeric(unlist(p_value)))
  }else {
    pvalue=test%>%
      mutate(p_value=tryCatch({t.test(unlist(Dep_Values_No_Treatment), unlist(Dep_Values_Treatment))$p.value},
                              error=function(e) {NA}))
  }
  output=pvalue%>%
    select(-Dep_Values_No_Treatment,-Dep_Values_Treatment)%>%
    mutate(Signif=ifelse(p_value<.001, "***",ifelse(p_value<.01, "**",
                                                    ifelse(p_value<.05, "*", ifelse(p_value<.1, ".", "")))))%>%
    mutate(Factor=Factor)%>%
    mutate(Levels=as.character(Levels))%>%
    # mutate(across(one_of("No_Treatment", "Treatment", "Diff"), ~percent(.,accuracy=.1)))%>%
    relocate(Factor)
  
}

#' Produce the Pairwise Comparisons found Significant in a Report
#'
#' @param Report Report object output from Report/appendInteraction
#' @param Data The original data used to make Report
#'
#' @return
#' @import tidyverse
#' @export
#'
#' @examples
pairwiseReport <- function(Report, Data) {
  Ind=colnames(Report)[match("Diff",colnames(Report))-1]
  w=match("Measure",colnames(Report))+0
  x=match("Signif", colnames(Report))+1
  y=ncol(Report)+0
  z=w+1
  CompPrime=Report%>%
    select(1:w,x:y)%>%
    pivot_longer(z:ncol(.),names_to = "Factor", values_to = "p_value" )%>%
    filter(p_value%in%c(".","*","**","***"))%>%
    mutate(Row=row_number())
  comps=NULL
  for (i in 1:nrow(CompPrime)) {
    a=CompPrime%>%
      filter(Row==i)%>%
      select(-p_value, -Row)%>%
      pivot_longer(1:Factor,names_to = "Settings", values_to = "value")
    b=as.character(a$value)
    names(b)=a$Settings
    # c=Pairwise(data, "Treatment",b["Measure"],b["Factor"], b["Program"],b["Cohort"])
    c=Data%>%
      filter(if(b["Cohort"]!="All"){Cohort==b["Cohort"]} else {T})%>%
      filter(!is.na(!!b["Measure"]))%>%
      mutate_("Dependant"=b["Measure"], "Independant"=Ind, "Levels"=b["Factor"])%>%
      group_by(Independant, Levels)%>%
      summarize(Mean=mean(Dependant, na.rm=T),Count=n(), Measure=b["Measure"], Dep_Values=list(Dependant))%>%
      filter(Count>4)%>%
      select(-Count)%>%
      ungroup()%>%
      mutate(Independant=ifelse(Independant==1,"Treatment","No_Treatment"))%>%
      pivot_wider(names_from = Independant, values_from = c("Mean", "Dep_Values"))%>%
      rename(No_Treatment=Mean_No_Treatment, Treatment=Mean_Treatment)%>%
      filter(!is.na(No_Treatment)&!is.na(Treatment))%>%
      mutate(Diff=Treatment-No_Treatment)%>%
      group_by(Levels)%>%
      mutate(p_value=tryCatch({t.test(unlist(Dep_Values_No_Treatment), unlist(Dep_Values_Treatment))$p.value},
                              error=function(e) {NA}))%>%
      select(-Dep_Values_No_Treatment,-Dep_Values_Treatment)%>%
      mutate(Signif=ifelse(p_value<.001, "***",ifelse(p_value<.01, "**",
                                                      ifelse(p_value<.05, "*", ifelse(p_value<.1, ".", "")))))%>%
      mutate(Factor=b["Factor"], Cohort=b["Cohort"])
    comps[[i]]=c
  }
  comps=bind_rows(comps)%>%
    relocate(Cohort, Factor)%>%
    group_by(Cohort, Factor, Measure)%>%
    filter(min(p_value)<.05)%>%
    filter(max(row_number())>1)%>%
    ungroup()%>%
    mutate(Levels=str_replace_all(Levels, ":", " "),
           Factor=str_replace_all(Factor, ":"," X "))
  names(comps)[names(comps)=="Treatment"] <- Ind
  names(comps)[names(comps)=="No_Treatment"] <- "Matched_Cohort"
  comps
}


#' Creates a document including the Report and Pairwise comparisons produced above
#'
#' @param Report The report output by report/appendInteraction
#' @param Pairwise The output of the pairwise function created from Report
#' @param measureNames A named vector with cleaned names for Measures
#' @param factorNames A named vector with cleaned names for Factors
#' @param percentOutcomes A vector of Outcomes that should be percentages
#'
#' @return
#' @import tidyverse
#' @import officer
#' @import flextable
#' @export
#'
#' @examples
document <- function(Report, Pairwise, measureNames, factorNames, percentOutcomes) {
  x=match("Measure", colnames(Report))+1
  y=x+2
  z=x+4
  Report=Report%>%
    rename_with(~str_replace_all(.,"_"," "))
  Main=Report%>%
    filter(Cohort=="All")%>%
    mutate(across(x:y,~ifelse(Measure%in%percentOutcomes,percent(., accuracy=.1), round(.,3))))%>%
    mutate(Diff=ifelse(Diff>0, paste0("+",Diff), Diff))%>%
    mutate(Measure=fct_recode(Measure, !!!measureNames))%>%
    select(1:z, -`p value`, -Cohort, -Program)%>%
    flextable()%>%
    theme_vanilla()%>%
    bold(j=~Diff)%>%
    color(j=~Diff, color="Black")%>%
    color(~str_detect(Diff,"\\+"), ~Diff,color="#7C997C" )%>%
    color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"\\+"), ~Diff, color="#0db14b")%>%
    color(~str_detect(Diff,"-"), ~Diff,color="#93736C")%>%
    color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"-"), ~Diff, color="Red")%>%
    void(~Signif, part='all')%>%
    FitFlextableToPage()
  
  Cohorts=Report%>%
    filter(Cohort!="All")%>%
    mutate(across(x:y,~ifelse(Measure%in%percentOutcomes,percent(., accuracy=.1), round(.,3))))%>%
    mutate(Diff=ifelse(Diff==0,Diff,ifelse(Diff>0, paste0("+",Diff),Diff)))%>%
    mutate(Measure=fct_recode(Measure, !!!measureNames))%>%
    select(1:z, -`p value`, -Program)%>%
    flextable()%>%
    merge_v(j=~Cohort)%>%
    theme_vanilla()%>%
    bold(j=~Diff)%>%
    color(j=~Diff, color="Black")%>%
    color(~str_detect(Diff,"\\+"), ~Diff,color="#7C997C" )%>%
    color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"\\+"), ~Diff, color="#0db14b")%>%
    color(~str_detect(Diff,"-"), ~Diff,color="#93736C")%>%
    color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"-"), ~Diff, color="Red")%>%
    void(~Signif, part='all')%>%
    FitFlextableToPage()
  
  doc=read_docx()%>%
    body_add_flextable(Main)%>%
    body_add_par("Break")%>%
    body_add_flextable(Cohorts)
  
  for (k in levels(as.factor(Pairwise$Factor))) {
    Pairwise_Report=NULL
    x=match("Measure", colnames(Pairwise))+1
    y=x+2
    z=x+4
    Pairwise_Report=Pairwise%>%
      rename_with(~str_replace_all(.,"_"," "))%>%
      mutate(across(x:y,~ifelse(Measure%in%percentOutcomes,percent(., accuracy=.1), round(.,3))))%>%
      mutate(Diff=ifelse(Diff>0, paste0("+",Diff), Diff))%>%
      mutate(Measure=fct_recode(Measure, !!!measureNames))%>%
      filter(Factor==k)%>%
      select(-Factor, -`p value`)%>%
      rename(!!k:=Levels)%>%
      arrange(Cohort, Measure)%>%
      relocate(Cohort, Measure)%>%
      flextable()%>%
      merge_v(j=~Cohort)%>%
      merge_v(j=~Measure)%>%
      theme_vanilla()%>%
      bold(j=~Diff)%>%
      color(j=~Diff, color="Black")%>%
      color(~str_detect(Diff,"\\+"), ~Diff,color="#7C997C" )%>%
      color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"\\+"), ~Diff, color="#0db14b")%>%
      color(~str_detect(Diff,"-"), ~Diff,color="#93736C")%>%
      color(~Signif%in%c(".","*","**", "***")&str_detect(Diff,"-"), ~Diff, color="Red")%>%
      void(~Signif, part='all')%>%
      FitFlextableToPage()
    
    doc=body_add_par(doc,"Break")%>%
      body_add_flextable(Pairwise_Report)
  }
  doc
}




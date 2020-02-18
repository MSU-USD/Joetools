#' Standard Error
#'
#' @param data The vector or dataset over which the standard error is calculated
#'
#' @return The calulated Standard Error
#' @importFrom stats sd
#' @export
#'
#' @examples
#' x=c(10, 20, 30, 40, 50)
#' se(x)
se=function(data){
  return(sd(data, na.rm=T)/sqrt(sum(!is.na(data))))
}

#' Flexible Workbook Creation Function
#'
#' @param df The dataframe to be saved
#' @param filename Name of the file, including the xlsx extension
#' @param sheetBy Option to create sheets based on the values in a particular column of the dataframe
#' @param keepNames The function automatically renames the column names with \code{\link[stringr:case]{stringr::str_to_sentence()}}. Setting this to TRUE will leave your column names unaltered
#'
#' @return Creates a Excel workbook from the supplied dataframe
#' @import tidyverse
#' @import xlsx
#' @import lazyeval
#' @export
#'
#' @examples
#' library(datasets)
#' library(tidyverse)
#' library(xlsx)
#' data(iris)
#' wbsave(iris, "Iris report.xlsx")
#' wbsave(iris, "Iris report - by Species.xlsx", sheetBy="Species", keepNames=TRUE)
wbsave=function(df,filename,sheetBy=NULL,keepNames=FALSE){
  if(keepNames==FALSE) {names(df)=str_to_sentence(names(df))}
  if(!is.null(sheetBy)) {sheetBy=str_to_sentence(sheetBy)}
  wb=createWorkbook()
  TABLE_COLNAMES_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=11, color="#44546A", isBold=TRUE)+
    Border(color = "#8EA9DB", position = "BOTTOM", pen="BORDER_THICK")
  sheetAll=createSheet(wb,sheetName = "All")
  wbdata=NULL
  wbAll=NULL
  if(!is.null(sheetBy)) {l=levels(as.factor(df[[sheetBy]]))}
  if(!is.null(sheetBy)) {for (i in l){
    filter_criteria <- interp(~y == x, .values=list(y = as.name(sheetBy), x = i))
    wbdata=df%>%
      filter_(filter_criteria)
    wbAll[[i]]=wbdata
    sheet=createSheet(wb,sheetName = i)
    addDataFrame(wbdata,sheet, colnamesStyle = TABLE_COLNAMES_STYLE)
    autoSizeColumn(sheet, colIndex = 1:ncol(df)+1)
  }}
  addDataFrame(df,sheetAll, colnamesStyle = TABLE_COLNAMES_STYLE)
  autoSizeColumn(sheetAll, colIndex = 1:ncol(df)+1)
  saveWorkbook(wb, filename)
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

report=function(df, Measures, Factor, paired=c("Yes", "No", "Try")){
  output=df%>%
    group_by(.data[[Factor]])%>%
    summarise_at(vars(Measures), mean,na.rm=T)%>%
    mutate(Levels =ifelse(.data[[Factor]]==1, "Treatment", "No_Treatment"))%>%
    select(-.data[[Factor]])%>%
    pivot_longer(cols=Outcomes, names_to = "Measure")%>%
    pivot_wider(names_from = "Levels", values_from = "value")%>%
    mutate(Diff=Treatment-No_Treatment)%>%
    mutate(Measure=factor(.data$Measure,levels=Measures))%>%
    arrange(Measure)


  reportpvaluecount=0
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

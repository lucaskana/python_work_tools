# install.packages("plan")
# https://cran.r-project.org/web/packages/plan/vignettes/plan.html
#install.packages("languageserver")

library("tidyverse")
library("plan")
library("readxl")
# GGPLOT GANTT CHART
library(ggplot2)
library(tidyr)
library(lubridate)

title <- "Title"
project <- 'project/'
datafolder <- './data/cronograma/'
datafolder <- paste(datafolder,project, sep = '')
current_date <- '2023-05-24'
cronograma <- paste(datafolder,"cronograma_R.xlsx", sep = '')
cronograma_png <- paste(datafolder,'timeline.png', sep = '')

df_chart <- read_excel(cronograma, sheet = "CronogramaMacro")
#View(df_chart)

df_chart <- df_chart %>% 
  mutate(Task = coalesce(Task,Atividades))

g <- new("gantt")

for(i in 1:24)
{
  if ((i %% 2) != 0) {
    completetion <- 100
  } else {
    completetion <- 0
  }
  print(completetion)
  g <- ganttAddTask(g, df_chart[["Task"]][i], as.character(df_chart[["inicio"]][i]), as.character(df_chart[["fim"]][i]), done=completetion)
}

png(cronograma_png, width = 900, height = 480)
plot(g,event.label='Data atual',event.time=current_date,
     col.event=c("red"),
     col.done=c("green3"),
     col.notdone=c("orange2"),
     main=title
)
legend("topright", pch=22, pt.cex=2, cex=0.9, pt.bg=c("orange2", "green3"),
       border="black",
       legend=c("Planejado", "Executado"), bg="white", xpd=TRUE)
dev.off() # to complete the writing process and return output to your monitor
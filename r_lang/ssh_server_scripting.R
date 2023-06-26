install.packages("ssh")
install.packages("properties")
install.packages("clipr")

library(ssh)
library(properties)
library("clipr")

#####################################################################################
# 
# 1 - Le properties file de ./data/ssh_carga.properties
#    ssh_host=user@host
#    password=senha
#    listcmd=ls - arquivos para fazer download
#
# 2 - Abre SSH e executa comandos de data/command.RData
#
# 3 - Download para pasta destino_folder
#
# 4 - Cria zip da pasta destino_folder
#
#####################################################################################

# 1 - Le properties file de ./data/ssh_carga.properties
prop <- read.properties(
  './data/ssh_carga.properties', 
  fields = c("ssh_host", "password","listcmd")
  )
write_clip(prop$password) 
session <- ssh_connect(prop$ssh_host, verbose = 2)

load("data/command.RData")
View(df)
#save(df, file = "data/command.RData")

# 2 - Abre SSH e executa comandos de data/command.RData
for (p in df['command']) {
  print(p)
  out <- ssh_exec_wait(session, command = p)
}

commandline <- prop$listcmd
results <- capture.output(ssh_exec_wait(session, command = commandline))
results

# 3 - Download para pasta destino_folder
destino_folder <- "./data"
for (p in results) {
  if(!is.na(strsplit(p, "\\s+")[[1]][11])){
    print(strsplit(p, "\\s+")[[1]][11])
    filename <- strsplit(p, "\\s+")[[1]][11]
    scp_download(session, filename, to = destino_folder)
  }
}

ssh_disconnect(session)

# 4 - Cria zip da pasta destino_folder
zippath <- "C:\\PROGRA~1\\7-Zip\\7z"
zipcommand <- "a carga.zip ./data"
system(paste(zippath,zipcommand))

#####################################################################################
#
# SandBox
#
#####################################################################################
#
# write.properties(file = './data/ssh_carga.properties',
#    properties = list(
#      ssh_host = "user@localhost", 
#      password = "password",
#      listcmd = 'find  -mtime -1 -ls'
#      ),
#    fields = c("ssh_host", "password","listcmd")
#    )
#
# read.properties('./data/ssh_carga.properties', fields = c("ssh_host", "password"))
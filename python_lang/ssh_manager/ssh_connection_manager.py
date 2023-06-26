import paramiko
import os

############################################################################
#
# Conecta via ssh em um servidor
#
############################################################################
class ConnectionManager:
    def __init__(self,hostname,username,password='', port=22, destinationfolder='.') -> None:
        self.hostname = hostname
        self.username = username
        self.password = password
        self.port = port
        self.destinationfolder = destinationfolder
    
    ############################################################################
    #
    # Executa lista de comandos
    #
    ############################################################################
    def executeCommandList(self, commandlist):
        print(f'Executando {commandlist}')
        result = []
        with paramiko.SSHClient() as client:

            client.load_system_host_keys()
            client.connect(self.hostname, self.port, self.username)

            for command in commandlist:
                (stdin, stdout, stderr) = client.exec_command(command)
                output = stdout.readlines()
                result.append(output)
        return [item for sublist in result for item in sublist]

    ############################################################################
    #
    # Realiza Download Files
    #
    ############################################################################
    def downloadFiles(self, filesList) -> None:
        with paramiko.SSHClient() as client:
            client.load_system_host_keys()
            client.connect(self.hostname, self.port, self.username)

            with client.open_sftp() as sftp:
                for file in filesList:
                    filename = file.split('/')[-1]
                    print("Remote : <" + file + "> Local: <" f"{self.destinationfolder}/{filename}" + ">")
                    sftp.get(file.replace('\n',''), f"{self.destinationfolder}/{filename}".replace('\n',''))

    ############################################################################
    #
    # Limpar Local Folder
    #
    ############################################################################
    def cleanDestinationFolder(self) -> None:
        if("." == self.destinationfolder):
            print("Define destination folder - " + self.destinationfolder)
        else:
            print("Cleaning destination folder" + self.destinationfolder)
            [os.remove(os.path.join(self.destinationfolder, file)) for file in os.listdir(self.destinationfolder)]

    def __str__(self) -> str:
        return f'The hostname is {self.hostname} and the username is {self.username}'
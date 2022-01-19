Dim WshShell = CreateObject("WScript.Shell") 'declara a variavel WshSheell 
On Error Resume Next 'continua executando mesmo se der erro
WshShell.Run("CMD /K query session |clip") 'chama o CMD buscando a Query Sessio e copiando para o clipboard
WshShell.Run("taskkill /f /im cmd.exe") 'mata o CMD
On Error Goto 0 'Volta o estado normal da execução (volta a olhar para os erros)

Dim Clipboard = CreateObject("htmlfile").ParentWindow.ClipboardData.getData("text") 'pega os dados que estão no clipboard e traz para variavel clipboard

Dim numero 'nova variavel numero
if InStr(Clipboard,Environment.UserName) > 0 'procura se no texto do clipboard existe o nome de usuario da maquina.
numero = InStr(Clipboard,Environment.UserName) 'coloca o numero do caractere aonde encontrou o nome da maquina

For i = numero to 99999 ' faz um loop
i = i + 1 
if Clipboard(i) > "0" And Clipboard(i) < "9" 'o proximo caractere é um numero entre zero e 9 ?

  NumeroSessao = Clipboard(i) 'caractere do clipboard adicionado na variavel NumeroSessao

  exit for 'sai do for
end if 'sai do if

Next 'proximo caractere
end if 'sai do if

TarefasChrome = New DataTable 'declara uma nova tabela de dados com o nome TarefasChrome
Dim objWMIService = GetObject("winmgmts:\\.\root\cimv2") 'declara a variavel o local da query procurada
Dim colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chrome.exe'") 'declara o filtro da Query
Dim objProcess ' declara uma nova variavel
Dim contagem = 0 'declara uma nova variavel
TarefasChrome.Columns.Add("Name", GetType(String)) 'na adiciona uma nova coluna na tabela tarefaschorme com o nome de "Name" e diz que ele é um texto
TarefasChrome.Columns.Add("PID", GetType(Integer)) 'na adiciona uma nova coluna na tabela tarefaschorme com o nome de "PID" e diz que ele é um numero inteiro
TarefasChrome.Columns.Add("NomeUsuario", GetType(String)) 'na adiciona uma nova coluna na tabela tarefaschorme com o nome de "NomeUsuario" e diz que ele é um texto
TarefasChrome.Columns.Add("ParentProcessId", GetType(Integer)) 'na adiciona uma nova coluna na tabela tarefaschorme com o nome de "ParentProcessId" e diz que ele é um texto
For Each objProcess In colProcesses 'para cada linha da query filtrada

if objProcess.SessionId = NumeroSessao then 'valida se o SessionId = NumeroSessão

If objProcess.GetOwner = 0 Then 'valida se é o primeiro item encontrado?
contagem = contagem + 1 '"adiciona 1 na contagem para que pegue somente o primeira linha da query"
if contagem = 1 ' se contagem for igual a 1
  PIDGoogleChrome = objProcess.ProcessID 'pega o numero de ID
end if 'sai do if
TarefasChrome.Rows.Add(New Object() {objProcess.Name, objProcess.ProcessID,Environment.UserName, objProcess.ParentProcessId}) 'adiciona na collection tudo que pegou na query
'TarefasChrome.Rows.Add(objProcess.Name)
end if
end if
Next

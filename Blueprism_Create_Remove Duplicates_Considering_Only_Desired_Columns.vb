'Necessário obter "System.Collections.Generic" em inicialize.
'Colections necessárias input: [Collection In], [Field Names] output: [Collection Out]
Dim Tabela = New DataTable
Dim Tabela2 = New DataTable

Dim listOfFields = New List(Of String)()
Dim Nomes_Coluna = New List(Of String)()

For Each row As System.Data.DataRow In Field_Names.Rows

	listOfFields.Add(row(0).ToString())

Next

Tabela = Collection_In.DefaultView.ToTable(True,listOfFields.ToArray()) 'Tabela com as duplicatas removidas observando somente as colunas desejadas.
'Collection_Return = Collection_In.DefaultView.ToTable(True,listOfFields.ToArray()) 'Enviando o Mesmo para uma collection, só de conferencia.


Tabela2 = Tabela.clone() 'Criando um Clone de somente cabeçalho para tabela2

'transformando os valores da tabela + nome do cabeçaho, para criar um Texto_Filtro (para depois ser usado com commando de SQL)
For I As integer = 0 To Tabela.Rows.Count - 1
		Dim Values(Tabela.Columns.Count - 1) As Object
		For J As Integer = 0 To Tabela.Columns.Count - 1
			Values(J) =  Tabela.Columns(J).ColumnName &"='"& Tabela.Rows(I)(J)&"'"
		Next

		Tabela2.Rows.Add(Values) 'Transfere os valores da Tabela + Nome da Coluna Para tabela2.
	Next	

Collection_Out = Collection_In.Clone 'Clona o cabeçalho da Collection_in para Collection_Out

'Começa a criação de texto Filtrado.
Dim Texto_Filtrado = ""
Dim Linha_Fim = Tabela2.Rows.Count

Dim Linha_Contagem = 0
For Each V As DataRow In Tabela2.Rows
Texto_Filtrado = ""
Linha_Contagem = Linha_Contagem + 1
For J As Integer = 0 To Tabela2.Columns.Count - 1

If Linha_Contagem <= Linha_Fim And J < Tabela2.Columns.Count -1

Texto_Filtrado = Texto_Filtrado + V(J) &" AND "  'Cria o Filtro esperando a existencia de mais colunas
Else
Texto_Filtrado = Texto_Filtrado + V(J) 'Cria o Filtro sem continuidade.
End If
Next
' termina a criação de texto filtrado e inicia a captura do dado na celula exata com filtro.

Dim Values(Collection_In.Columns.Count - 1) As Object
For J As Integer = 0 To Collection_In.Columns.Count - 1
			Values(J) = Collection_In.Select(Texto_Filtrado)(0)(J)
		Next
		Collection_Out.Rows.Add(Values)

Next




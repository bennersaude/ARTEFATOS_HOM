'HASH: 55ACA37A9CAEE177F690867CBE582938
'Macro: SIS_CONTABCAMPOS

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	CanContinue = True

	Dim SQL1 As BPesquisa
    Set SQL1 = NewQuery

    SQL1.Add("SELECT HANDLE                          ")
    SQL1.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO      ")
    SQL1.Add(" WHERE CAMPOINTEGRACAOCOMPRAS = :HANDLE OR ")
	SQL1.Add("       CAMPOMODULOBENEFICIARIO = :HANDLE OR ")
	SQL1.Add("       CAMPOCORPORATIVO = :HANDLE OR   ")
	SQL1.Add("       CAMPOTESOURARIA = :HANDLE OR    ")
	SQL1.Add("       CAMPOCARENCIA = :HANDLE OR      ")
	SQL1.Add("       CAMPOREAJUSTES = :HANDLE OR     ")
	SQL1.Add("       CAMPOFATURALANC = :HANDLE OR    ")
	SQL1.Add("       CAMPOFATURA = :HANDLE OR        ")
	SQL1.Add("       CAMPODEMONSTRATIVOUTILIZACAO = :HANDLE OR ")
	SQL1.Add("       CAMPOAUTORIZACAO = :HANDLE OR   ")
	SQL1.Add("       CAMPONOTAFISCAL = :HANDLE OR    ")
	SQL1.Add("       CAMPOMBENEFICIARIO = :HANDLE OR ")
	SQL1.Add("       CAMPOSISTEMA = :HANDLE OR       ")
	SQL1.Add("       CAMPOCONTABILIDADE = :HANDLE OR ")
	SQL1.Add("       CAMPO = :HANDLE OR              ")
	SQL1.Add("       CAMPOCARTAO = :HANDLE OR        ")
	SQL1.Add("       ORDEM= :HANDLE                  ")
    SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL1.Active = True

	If SQL1.FieldByName("HANDLE").AsInteger > 0 Then
		bsShowMessage("Campo faz parte de algum modelo de leiaute e não pode ser excluído! Retire-o do modelo e tente novamente!", "E")
		CanContinue = False
    	Set SQL1 = Nothing
		Exit Sub
	End If

    Set SQL1 = Nothing
End Sub

Public Sub TABLE_NewRecord()

  Select Case NodeInternalCode
    Case 1
      CurrentQuery.FieldByName("ORIGEM").Value = "1"
      CurrentQuery.FieldByName("NOME").Value = "CFI."

    Case 2
      CurrentQuery.FieldByName("ORIGEM").Value = "2"
      CurrentQuery.FieldByName("NOME").Value = "FAT."

    Case 3
      CurrentQuery.FieldByName("ORIGEM").Value = "3"
      CurrentQuery.FieldByName("NOME").Value = "LFA."

    Case 4
      CurrentQuery.FieldByName("ORIGEM").Value = "4"
      CurrentQuery.FieldByName("NOME").Value = "DOC."

    Case 5
      CurrentQuery.FieldByName("ORIGEM").Value = "5"
      CurrentQuery.FieldByName("NOME").Value = "PAR."

    Case 6
      CurrentQuery.FieldByName("ORIGEM").Value = "6"
      CurrentQuery.FieldByName("NOME").Value = "LPA."

    Case 7
      CurrentQuery.FieldByName("ORIGEM").Value = "7"
      CurrentQuery.FieldByName("NOME").Value = "TES."

    Case 8
      CurrentQuery.FieldByName("ORIGEM").Value = "8"
      CurrentQuery.FieldByName("NOME").Value = "LTE."

    Case 9
      CurrentQuery.FieldByName("ORIGEM").Value = "S"
      CurrentQuery.FieldByName("NOME").Value = "SIS."

  End Select

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("ORIGEM").AsString = "I" And CurrentQuery.FieldByName("TABELA").AsInteger <= 0 Then
		bsShowMessage("O campo 'Tabela' é obrigatório!", "E")
		CanContinue = False
	End If

End Sub

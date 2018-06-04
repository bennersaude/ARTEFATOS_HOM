'HASH: 9CF4D1E12EB79C299C4A6D63FD9A6646
Option Explicit
Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("PERMITIRSOLICPORTSERVICO").AsString  = "S" Then
		If CurrentQuery.FieldByName("SEVERIDADE").IsNull Then
			MsgBox("Campo Severidade é obrigatório")
			CanContinue = False
		End If
	End If

	If CurrentQuery.FieldByName("OUVIDORIA").AsString = "1" Then
	  If CurrentQuery.FieldByName("PRAZOINICIALOUVIDORIA").IsNull Or CurrentQuery.FieldByName("PRAZOMAXIMOOUVIDORIA").IsNull Then
	    MsgBox("Todos os prazos de ouvidoria são obrigatórios!")
	    CanContinue = False
	  End If
	End If

	If PossuiSegmentoSetor Then
		MsgBox("Existem segmentos relacionados ao setor : "+RetornaDescricaoSetor+" !")
		CanContinue = False
	End If

End Sub
Public Function PossuiSegmentoSetor As Boolean

  PossuiSegmentoSetor = False
  Dim SQL As Object
  Set SQL=NewQuery
  SQL.Clear
  SQL.Add("SELECT SETOR FROM CA_CLASSEOUTROS WHERE HANDLE = :CLASSE")
  SQL.ParamByName("CLASSE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If CurrentQuery.FieldByName("SETOR").AsInteger <> SQL.FieldByName("SETOR").AsInteger Then

  	  SQL.Clear
	  SQL.Add("SELECT COUNT(*) QTDE FROM CA_CLASSEOUTROS_SEGMENTO WHERE CLASSEOUTROS = :CLASSE")
	  SQL.ParamByName("CLASSE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	  SQL.Active = True

	  If SQL.FieldByName("QTDE").AsInteger > 0 Then
	    PossuiSegmentoSetor = True
	  End If

  End If

  Set SQL = Nothing

End Function
Public Function RetornaDescricaoSetor As String

  Dim viSetor As Integer
  Dim SQL As Object
  Set SQL=NewQuery

  SQL.Clear
  SQL.Add("SELECT SETOR FROM CA_CLASSEOUTROS WHERE HANDLE = :CLASSE")
  SQL.ParamByName("CLASSE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  viSetor = SQL.FieldByName("SETOR").AsInteger

  SQL.Clear
  SQL.Add("SELECT DESCRICAO FROM SAM_SETOR WHERE HANDLE = :SETOR")
  SQL.ParamByName("SETOR").Value = viSetor
  SQL.Active = True

  RetornaDescricaoSetor = SQL.FieldByName("DESCRICAO").AsString

  Set SQL = Nothing

End Function

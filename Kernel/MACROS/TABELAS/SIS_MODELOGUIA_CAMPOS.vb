'HASH: 95028C120AE60D0BCB390B600EFB7919


Public Sub BOTAOESCOLHECAMPO_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterios As String
  Dim SQL As Object

  Set SQL = NewQuery

  Set interface = CreateBennerObject("Procura.Procurar")
  vColunas = "NOME"
  vCampos = "Nome ZCampo"

  If CurrentQuery.State = 1 Then
    CurrentQuery.Edit
  End If

  CurrentQuery.UpdateRecord
  If CurrentQuery.FieldByName("EVENTOGUIA").AsString = "G" Then
    vCriterios = "TABELA = (SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SAM_GUIA')"
  Else
    If CurrentQuery.FieldByName("EVENTOGUIA").AsString = "E" Then
       vCriterios = "TABELA = (SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SAM_GUIA_EVENTOS')"
    Else
       vCriterios = "TABELA = (SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SAM_GUIA_EVENTOS_MATMED')"
    End If
  End If
  vHandle = interface.Exec(CurrentSystem, "Z_CAMPOS", vColunas, 1, vCampos, vCriterios, "Campos da tabela de guias e eventos", True, "")

  SQL.Add "SELECT HANDLE, NOME FROM Z_CAMPOS WHERE HANDLE = :HANDLE"
  SQL.ParamByName("HANDLE").Value = vHandle
  SQL.Active = True

  CurrentQuery.FieldByName("ZCAMPO").Value = SQL.FieldByName("NOME").AsString
  CurrentQuery.FieldByName("LEGENDA").Value = UCase(Mid(CurrentQuery.FieldByName("ZCAMPO").AsString, 1, 1)) + LCase(Mid(CurrentQuery.FieldByName("ZCAMPO").AsString, 2))

  Set interface = Nothing
  Set SQL = Nothing

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		DEPENDE.WebLocalWhere = "EVENTOGUIA = @CAMPO(EVENTOGUIA)"
	ElseIf VisibleMode Then
		DEPENDE.LocalWhere = "EVENTOGUIA = @EVENTOGUIA"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		DEPENDE.WebLocalWhere = "EVENTOGUIA = @CAMPO(EVENTOGUIA)"
	ElseIf VisibleMode Then
		DEPENDE.LocalWhere = "EVENTOGUIA = @EVENTOGUIA"
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOESCOLHECAMPO" Then
		BOTAOESCOLHECAMPO_OnClick
	End If
End Sub

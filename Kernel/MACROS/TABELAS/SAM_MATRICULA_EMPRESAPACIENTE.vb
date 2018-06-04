'HASH: 8668213FD5F3C4BEE0CCB880EF26098D


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim interface As Object
  Dim Linha As String
  Dim CAMPO As String
  Dim CONDICAO As String
  Dim SQL As Object
  Set SQL = NewQuery

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = interface.Vigencia(CurrentSystem, "SAM_MATRICULA_EMPRESAPACIENTE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MATRICULA", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    MsgBox(Linha)
  End If
  Set interface = Nothing

  SQL.Clear
  SQL.Add(" SELECT SITUACAO FROM SAM_EMPRESAPACIENTE WHERE HANDLE = :EMPRESA ")
  SQL.ParamByName("EMPRESA").AsInteger = CurrentQuery.FieldByName("EMPRESAPACIENTE").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "I" Then
    MsgBox(" A empresa está INATIVA ")
    CanContinue = False
  End If

End Sub


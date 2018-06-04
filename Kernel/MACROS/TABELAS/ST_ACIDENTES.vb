'HASH: 07106CE95EF5C116C8909BBF9326F198
'MACRO: ST_ACIDENTES

Public Sub DATAHORA_OnChange()
  Dia = Weekday(CurrentQuery.FieldByName("DATAHORA").AsDateTime)
  CurrentQuery.FieldByName("DIASEMANA").AsInteger = Dia
End Sub

Public Sub IMPRIMIR_OnClick()

  Set Sql = NewQuery

  Sql.Add ("SELECT  HANDLE FROM R_RELATORIOS  WHERE (CODIGO = :PCODIGORELATORIO)")
  Sql.ParamByName("PCODIGORELATORIO").AsString = "PCMSO004"
  Sql.Active = True
  If Not Sql.EOF Then
    ReportPreview(Sql.FieldByName("HANDLE").AsInteger, "", F, F)
  End If
  Set Sql = Nothing

End Sub

Public Sub PACIENTE_OnPopup(ShowPopup As Boolean)
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampo As String
  Dim vTabelas As String
  Dim vHandle As Integer
  Dim Interface As Object
  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_BENEFICIARIO.BENEFICIARIO|SAM_MATRICULA.NOME|SAM_MATRICULA.MATRICULA|MS_PACIENTES.IDADE"
  vCriterio = "MS_PACIENTES.FILIAL = " + Str(CurrentBranch)
  vCampos = "Beneficiário|Nome|Matrícula|Idade"
  vTabelas = "MS_PACIENTES|SAM_BENEFICIARIO[SAM_BENEFICIARIO.HANDLE=MS_PACIENTES.BENEFICIARIO]|SAM_MATRICULA[SAM_MATRICULA.HANDLE=MS_PACIENTES.MATRICULA]"

  vHandle = Interface.Exec(CurrentSystem, vTabelas, vColunas, 2, vCampos, vCriterio, "Paciente", False, PACIENTE.Text)

  If vHandle > 0 Then
    CurrentQuery.FieldByName("PACIENTE").AsInteger = vHandle
    Dim Sql As Object, SqlAux As Object
    If VisibleMode Then
      TextoSql = "SELECT * FROM ST_ACIDENTES  WHERE "
      TextoSql = TextoSql + " PACIENTE = " + CurrentQuery.FieldByName("PACIENTE").AsString
      Set SqlAux = NewQuery
      SqlAux.Add(TextoSql)
      SqlAux.Active = True

      If Not SqlAux.EOF Then
        CurrentQuery.FieldByName("ACIDENTEANTES").Value = "S"
      Else
        CurrentQuery.FieldByName("ACIDENTEANTES").Value = "N"
      End If
    End If
  End If
End Sub

Public Sub TABLE_AfterInsert()
  Dia = Weekday(CurrentQuery.FieldByName("DATAHORA").AsDateTime)
  CurrentQuery.FieldByName("DIASEMANA").AsInteger = Dia
End Sub


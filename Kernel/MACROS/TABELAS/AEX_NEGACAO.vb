'HASH: 422703939FEA3163D8C029F3E7EB6D8E
' atualizada em 10/08/2007
 '#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("CODAUXNEGACAOSISTEMA").AsInteger = 0
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' ******** INICIO SMS - 39741 - 01/07/2005 - DRUMMOND ********
  Dim QVerifica As Object
  Dim QVerificaCodEx As Object
  Dim QUpdate As Object

  Set QVerifica = NewQuery
  QVerifica.Clear
  QVerifica.Add("SELECT COUNT(1) QTD")
  QVerifica.Add("  FROM AEX_NEGACAO")
  QVerifica.Add(" WHERE EMPCONECT = :EMPCO")
  QVerifica.Add("   AND CODAUXNEGACAOSISTEMA = :CODAUX")
  QVerifica.Add("   AND HANDLE <> :HNDL")
  QVerifica.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  QVerifica.ParamByName("CODAUX").AsInteger = CurrentQuery.FieldByName("CODAUXNEGACAOSISTEMA").AsInteger
  QVerifica.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QVerifica.Active = True

  If QVerifica.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Negação já cadastrada.", "E")
    CanContinue = False
  End If

  '	Set QVerificaCodEx = NewQuery
  '	QVerificaCodEx.Clear
  '	QVerificaCodEx.Add("SELECT COUNT(1) QTD")
  '	QVerificaCodEx.Add("  FROM AEX_NEGACAO")
  '	QVerificaCodEx.Add(" WHERE CODNEGACAOEXTERNO = :CODEX")
  '	QVerificaCodEx.Add("   AND EMPCONECT = :EMPCO")
  '	QVerificaCodEx.ParamByName("CODEX").AsInteger = CurrentQuery.FieldByName("CODNEGACAOEXTERNO").AsInteger
  '	QVerificaCodEx.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  '	QVerificaCodEx.Active = True

  '    CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

  '	If (CanContinue) And (QVerificaCodEx.FieldByName("QTD").AsInteger > 0) Then
  '		MsgBox "Código externo digitado já existe."
  '		CanContinue = False
  '	End If
  '	Set QVerificaCodEx = Nothing

  '	Set QVerifica = Nothing

  'SMS 87108 - Adequação do sistema Web
  Dim QDescricao As Object
  Dim SQL As Object
  Dim vCodigo As Integer

  Set SQL = NewQuery
  Set QDescricao = NewQuery
  QDescricao.Clear
  QDescricao.Add("SELECT DESCRICAO,")
  QDescricao.Add("       CODIGO")
  QDescricao.Add("  FROM SIS_MOTIVONEGACAO")
  QDescricao.Add(" WHERE HANDLE = :HNDL")
  QDescricao.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("CODNEGACAOSISTEMA").AsInteger
  QDescricao.Active = True

   CurrentQuery.FieldByName("CODAUXNEGACAOSISTEMA").AsInteger = QDescricao.FieldByName("CODIGO").AsInteger


  If QDescricao.FieldByName("DESCRICAO").AsString <> "" Then
   CurrentQuery.FieldByName("DESCRICAOSISTEMA").AsString = QDescricao.FieldByName("DESCRICAO").AsString
  End If
 Set QDescricao = Nothing





  If CanContinue Then
    Set QUpdate = NewQuery

    QUpdate.Clear
    QUpdate.Add("UPDATE AEX_NEGACAO SET")
    QUpdate.Add("       PROCESSADO = 'N'")
    QUpdate.Add(" WHERE EMPCONECT = :EMPCONECT")
    QUpdate.ParamByName("EMPCONECT").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
    QUpdate.ExecSQL

    Set QUpdate = Nothing
  End If
  ' ******** FIM    SMS - 39741 - 01/07/2005 - DRUMMOND ********
End Sub

Public Sub TABLE_NewRecord()
  ' ******** INICIO SMS - 39741 - 01/07/2005 - DRUMMOND ********
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
  ' ******** FIM    SMS - 39741 - 01/07/2005 - DRUMMOND ********
End Sub


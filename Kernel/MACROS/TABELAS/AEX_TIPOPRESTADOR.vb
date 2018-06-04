'HASH: 4905EBDE0B1175954135A21494A9738C
' atualizada em 10/08/2007
'#Uses "*bsShowMessage"



Public Sub CODTIPOSISTEMA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCabecs As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim vTitulo As String

  If CODTIPOSISTEMA.PopupCase <>0 Then
    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vCabecs = "Código|Descrição"
    vCriterio = " CODIGO NOT IN (SELECT CODTIPOSISTEMA FROM AEX_TIPOPRESTADOR) "
    vColunas = "CODIGO|DESCRICAO"
    vTabela = "SAM_TIPOPRESTADOR"
    vTitulo = "Tipo de prestadores"

    vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("CODTIPOSISTEMA").AsInteger = vHandle
    End If
    Set interface = Nothing
  Else
    ShowPopup = True
  End If


End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim QUpdate As Object

  Set SQL = NewQuery
  SQL.Add("SELECT DESCRICAO FROM SAM_TIPOPRESTADOR WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CODTIPOSISTEMA").AsInteger
  SQL.Active = True
  CurrentQuery.FieldByName("DESCRICAOSISTEMA").AsString = SQL.FieldByName("DESCRICAO").AsString

  ' Set SQL1 = NewQuery
  ' SQL1.Clear
  ' SQL1.Add("SELECT COUNT(1) QTD")
  ' SQL1.Add("  FROM AEX_TIPOPRESTADOR")
  ' SQL1.Add(" WHERE EMPCONECT = :EMPCO")
  ' SQL1.Add("   AND CODTIPOEXTERNO = :CODEX")
  ' SQL1.Add("   AND HANDLE <> :HNDL")
  ' SQL1.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  ' SQL1.ParamByName("CODEX").AsInteger = CurrentQuery.FieldByName("CODTIPOEXTERNO").AsInteger
  '  SQL1.ParamByName("HNDL").AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger
  '  SQL1.Active = True

  ' If SQL1.FieldByName("QTD").AsInteger > 0 Then
  '   MsgBox "Código externo digitado já existe."
  '  	CanContinue = False
  '  End If

  '  Set SQL1 = Nothing

  Set SQL2 = NewQuery
  
  SQL2.Clear
  SQL2.Add("SELECT COUNT(1) QTD")
  SQL2.Add("  FROM AEX_TIPOPRESTADOR")
  SQL2.Add(" WHERE EMPCONECT = :EMPCO")
  SQL2.Add("   And CODTIPOSISTEMA = :CODAUX")
  SQL2.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  SQL2.ParamByName("CODAUX").AsInteger = CurrentQuery.FieldByName("CODTIPOSISTEMA").AsInteger
  SQL2.Active = True

  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

  If (CanContinue) And (SQL2.FieldByName("QTD").AsInteger > 1) Then
    bsShowMessage("O tipo de prestador já está existe.", "E")
    CanContinue = False
  End If
  Set SQL2 = Nothing

  If CanContinue Then
    Set QUpdate = NewQuery

    QUpdate.Clear
    QUpdate.Add("UPDATE AEX_TIPOPRESTADOR SET")
    QUpdate.Add("       PROCESSADO = 'N'")
    QUpdate.Add(" WHERE EMPCONECT = :EMPCO")
    QUpdate.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
    QUpdate.ExecSQL

    Set QUpdate = Nothing
  End If
End Sub


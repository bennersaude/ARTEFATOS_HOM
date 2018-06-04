'HASH: 8E9636DFC8FCC53FBF76687F7CCC924C
  '#Uses "*bsShowMessage"


Public Sub CODIGOSISTEMA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCabecs As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim vTitulo As String

  If CODIGOSISTEMA.PopupCase <>0 Then
    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vCabecs = "Código|Descrição"
    vColunas = "CODIGO|DESCRICAO"
    vTabela = "SAM_ESPECIALIDADE"
    vCriterio = " HANDLE NOT IN (SELECT CODIGOSISTEMA FROM AEX_ESPECIALIDADE)"
    vTitulo = "Especialidades"

    vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("CODIGOSISTEMA").AsInteger = vHandle
    End If
    Set interface = Nothing
  Else
    ShowPopup = True
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  '************ Inicio SMS - 39471 28/06/2005 Drummond ************
  Dim QueryCodEx As Object
  Dim SQL As Object
  Dim SQL1 As Object
  Dim QueryMudaFlag As Object

  Set QueryCodEx = NewQuery

  'Verifica no banco a existencia do código externo digitado
  QueryCodEx.Clear
  QueryCodEx.Add("SELECT COUNT(1) QTD FROM AEX_ESPECIALIDADE")
  QueryCodEx.Add(" WHERE CODIGOSISTEMA = :CODIGOSISTEMA     ")
  QueryCodEx.Add("   AND EMPCONECT = :EMPCO                 ")
  QueryCodEx.Add("   AND HANDLE <> :HANDLE                  ")
  QueryCodEx.ParamByName("CODIGOSISTEMA").AsInteger = CurrentQuery.FieldByName("CODIGOSISTEMA").AsInteger
  QueryCodEx.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  QueryCodEx.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QueryCodEx.Active = True

  If QueryCodEx.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Especialidade já existente para esta empresa de conectividade.", "E")
    CanContinue = False
    Exit Sub
  End If

  Set QueryCodEx = Nothing

  ' 	Set SQL1 = NewQuery

  ' 	SQL1.Clear
  '	SQL1.Add("SELECT COUNT(1) QTD FROM AEX_ESPECIALIDADE")
  '	SQL1.Add(" WHERE CODIGOEXTERNO = :CODEXT            ")
  '	SQL1.Add("   AND EMPCONECT = :EMPCO                 ")
  '	SQL1.Add("   AND HANDLE <> :HNDL                     ")
  '	SQL1.ParamByName("CODEXT").AsInteger = CurrentQuery.FieldByName("CODIGOEXTERNO").AsInteger
  '	SQL1.ParamByName("EMPCO").AsInteger  = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  '	SQL1.ParamByName("HNDL").AsInteger   = CurrentQuery.FieldByName("HANDLE").AsInteger
  '	SQL1.Active = True

  '	If SQL1.FieldByName("QTD").AsInteger > 0 Then
  '		MsgBox "O código externo digitado já existe."
  '		CanContinue = False
  '	End If

  '	Set SQL1 = Nothing

  Set SQL = NewQuery
  SQL.Add("SELECT DESCRICAO FROM SAM_ESPECIALIDADE WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CODIGOSISTEMA").AsInteger
  SQL.Active = True
  CurrentQuery.FieldByName("DESCRICAOSISTEMA").AsString = SQL.FieldByName("DESCRICAO").AsString
  Set SQL = Nothing

  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

  If CanContinue Then
    Set QueryMudaFlag = NewQuery

    'Muda o campo processado de todos os registros da tabela

    QueryMudaFlag.Clear
    QueryMudaFlag.Add("UPDATE AEX_ESPECIALIDADE SET PROCESSADO = 'N'")
    QueryMudaFlag.Add("       WHERE EMPCONECT = :EMPCO")
    QueryMudaFlag.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("EMPCONECT").AsInteger
    QueryMudaFlag.ExecSQL

    Set QueryMudaFlag = Nothing
  End If
  '************ FIM    SMS - 39471 28/06/2005 Drummond ************
End Sub

Public Sub TABLE_NewRecord()
  '************ Inicio SMS - 39471 28/06/2005 Drummond ************
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
  '************ Fim    SMS - 39471 28/06/2005 Drummond ************
End Sub


'HASH: 68FC208B593740A112553B6CD5E28767
'MACRO SAM_CHEQUE_CONSULTA
'#Uses "*bsShowMessage"
Dim HANDLECONTRATO As Integer


Public Sub BOTAOCANCELAR_OnClick()
  Dim interface As Object
  Dim vNum As Long
  Dim vNum2 As Long

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  vNum = CurrentQuery.FieldByName("NUMEROINICIAL").AsInteger
  vNum2 = CurrentQuery.FieldByName("NUMEROFINAL").AsInteger

  Set interface = CreateBennerObject("SamChequeConsulta.ChequeConsulta")
  interface.Cancelar(CurrentSystem, vNum, vNum2, CurrentUser)
  Set interface = Nothing
  BOTAOCANCELAR.Enabled = False
  BOTAOIMPRIMIR.Enabled = False

End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim interface As Object
  Dim pData As Date
  Dim pHandleContrato As Integer

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  pData = CurrentQuery.FieldByName("VALIDADE").AsDateTime
  pHandleContrato = CurrentQuery.FieldByName("CONTRATO").AsInteger
  Set interface = CreateBennerObject("SamChequeConsulta.ChequeConsulta")
  interface.Exec(CurrentSystem, pData, pHandleContrato)

  Set interface = Nothing
  BOTAOGERAR.Enabled = False
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim pHandleCheconsulta As Long

  pHandleCheconsulta = CurrentQuery.FieldByName("HANDLE").AsInteger

  Set interface = CreateBennerObject("SamChequeConsulta.ChequeConsulta")
  interface.Imprimir(CurrentSystem, pHandleCheconsulta)

  Set interface = Nothing

  ' Dim SQL As Object
  ' Set SQL =NewQuery
  ' SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CHE001'")
  ' SQL.Active =False
  ' SQL.Active =True

  'HandleRelatorio =SQL.FieldByName("HANDLE").AsInteger

  ' ReportPreview(HandleRelatorio,"A.CHECONSULTA="+CurrentQuery.FieldByName("HANDLE").AsString,True,False)
  ' Set SQL =Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object

  If Not CurrentQuery.FieldByName("NUMEROINICIAL").IsNull Then
    VALIDADE.ReadOnly = True
    BOTAOGERAR.Enabled = False
    BOTAOIMPRIMIR.Enabled = True
    BOTAOCANCELAR.Enabled = True
  Else
    VALIDADE.ReadOnly = False
    BOTAOGERAR.Enabled = True
    BOTAOIMPRIMIR.Enabled = False
    BOTAOCANCELAR.Enabled = False
  End If

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT COUNT(*) QTD FROM SAM_CONTRATO_CHECONSULTA_NUM WHERE CHECONSULTA=:CHECONSULTA AND CANCELADODATA IS NOT NULL")
  SQL.ParamByName("CHECONSULTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True


  If(SQL.FieldByName("QTD").AsInteger = (CurrentQuery.FieldByName("NUMEROFINAL").AsInteger - CurrentQuery.FieldByName("NUMEROINICIAL").AsInteger + 1))Then
  BOTAOCANCELAR.Enabled = False
  BOTAOIMPRIMIR.Enabled = False
Else
  If Not CurrentQuery.FieldByName("NUMEROINICIAL").IsNull Then
    BOTAOCANCELAR.Enabled = True
    BOTAOIMPRIMIR.Enabled = True
  End If
End If
Set SQL = Nothing

HANDLECONTRATO = CurrentQuery.FieldByName("CONTRATO").AsInteger


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :HANDLECONTRATO")
  SQL.ParamByName("HANDLECONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQL.Active = True

  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("O Contrato está cancelado, não é permitida a emissão de Cheque Consulta!!!", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Dim vData As Date
  Dim vDataHoje As Date

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT MAX(VALIDADE) MAXIMO FROM SAM_CONTRATO_CHECONSULTA WHERE CONTRATO=:CONTRATO AND VALIDADE>=:HOJE")
  SQL.ParamByName("CONTRATO").AsInteger = HANDLECONTRATO
  SQL.ParamByName("HOJE").AsDateTime = Date
  SQL.Active = True

  If(Not SQL.FieldByName("MAXIMO").IsNull)And(CurrentQuery.FieldByName("VALIDADE").IsNull)Then
  CurrentQuery.FieldByName("VALIDADE").AsDateTime = SQL.FieldByName("MAXIMO").AsDateTime
End If
If(SQL.FieldByName("MAXIMO").IsNull)And(CurrentQuery.FieldByName("VALIDADE").IsNull)Then

HANDLECONTRATO = RecordHandleOfTable("SAM_CONTRATO")


SQL.Clear
SQL.Add("SELECT DATAADESAO FROM SAM_CONTRATO WHERE HANDLE=:CONTRATO")
SQL.ParamByName("CONTRATO").AsInteger = HANDLECONTRATO

SQL.Active = True

If Not SQL.FieldByName("DATAADESAO").IsNull Then
  vData = SQL.FieldByName("DATAADESAO").AsDateTime
  vDataHoje = Date
  While(vData <vDataHoje)
  vData = DateAdd("yyyy", 1, vData)
Wend

CurrentQuery.FieldByName("VALIDADE").AsDateTime = vData
End If
End If

Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOIMPRIMIR"
			BOTAOIMPRIMIR_OnClick
	End Select
End Sub

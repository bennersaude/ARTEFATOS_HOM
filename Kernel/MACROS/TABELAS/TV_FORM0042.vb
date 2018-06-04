'HASH: ADA45D609DF329773D607169C49972BF
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
	Dim SFNCANCEL As Object
	Dim vsMensagem As String
	Dim viResult As Long
	Set SFNCANCEL = CreateBennerObject("SFNCANCEL.Cancelamento")

	viResult = SFNCANCEL.CancelaDocumento(CurrentSystem, _
										  CLng(SessionVar("HDOCUMENTO")), _
										  CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime, _
										  CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
										  CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsString, _
										  vsMensagem)

	Select Case viResult
		Case -1
			bsShowMessage("Processo abortado pelo usuário.", "I")
		Case 0
			bsShowMessage("Documento cancelado com sucesso.", "I")
		Case 1
			bsShowMessage(vsMensagem, "I")
	End Select
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Sql As BPesquisa
	Set Sql = NewQuery

	Sql.Clear
	Sql.Active = False
	Sql.Add("SELECT * ")
	Sql.Add("  FROM SFN_NOTA_DOCUMENTO ")
	Sql.Add(" WHERE DOCUMENTO = :DOCUMENTO ")
	Sql.ParamByName("DOCUMENTO").AsInteger = CLng(SessionVar("HDOCUMENTO"))
	Sql.Active = True

	If Not Sql.EOF Then
		bsShowMessage("O documento está vinculado a uma nota fiscal", "E")

		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim qTipoFatura As Object

  Set qTipoFatura = NewQuery
  qTipoFatura.Clear
  qTipoFatura.Add("Select Do.Handle")
  qTipoFatura.Add("  FROM SFN_DOCUMENTO        Do")
  qTipoFatura.Add("  Join SFN_DOCUMENTO_FATURA DF On DF.DOCUMENTO = Do.Handle")
  qTipoFatura.Add("  Join SFN_FATURA           FA On FA.Handle = DF.FATURA")
  qTipoFatura.Add("  Join SIS_TIPOFATURAMENTO  TF On TF.Handle = FA.TIPOFATURAMENTO")
  qTipoFatura.Add(" WHERE Do.Handle =" + SessionVar("HDOCUMENTO"))
  qTipoFatura.Add("   And TF.CODIGO = 500 ")
  qTipoFatura.Active = True

  If qTipoFatura.FieldByName("HANDLE").AsInteger > 0 Then
    bsShowMessage("Operação não permitida para um documento que possua pelo menos uma fatura de Provisão", "E")
    CanContinue = False
  End If

  Set qTipoFatura = Nothing
End Sub

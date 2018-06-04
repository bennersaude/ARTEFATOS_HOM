'HASH: 4B54F93AF988FFDBE7CE310E7266C98D
'#Uses "*bsShowMessage"

Public Sub BOTAOANEXARDOC_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Operação Cancelada. Registro está em Edição", "E")
  Else
  	Dim HandleAnexo As Long
  	Dim NomeArq, NomeArqServer As String

    NomeArq = OpenDialog
    NomeArqServer = NomeArq
    While InStr(NomeArqServer, "\")<>0
      NomeArqServer = Mid(NomeArqServer, InStr(NomeArqServer, "\") + 1,Len(NomeArqServer))
    Wend

    If NomeArq <>"" Then
      Dim QWork As BPesquisa
      Set QWork = NewQuery

      If Not InTransaction Then
        StartTransaction
      End If

      QWork.Clear
      QWork.Add("INSERT INTO SAM_INFORMATIVO_ANEXO (HANDLE, INFORMATIVO, DOCUMENTO, INCLUSAODATAHORA, INCLUSAOUSUARIO) ")
      QWork.Add("                           VALUES (:HANDLE, :INFORMATIVO,:DOCUMENTO,:DATAHORA,:USUARIO)")

      HandleAnexo = NewHandle("SAM_INFORMATIVO_ANEXO")
      QWork.ParamByName("HANDLE").AsInteger = HandleAnexo
      QWork.ParamByName("DATAHORA").AsDateTime = ServerNow
      QWork.ParamByName("DOCUMENTO").AsString = NomeArqServer
      QWork.ParamByName("USUARIO").AsInteger = CurrentUser
      QWork.ParamByName("INFORMATIVO").AsInteger = RecordHandleOfTable("SAM_INFORMATIVO")
      QWork.ExecSQL

      Dim Obj As Object
      Set Obj = SuperServerClient("DOC")
      Obj.Select("INFORMATIVOS\ANEXOS")
      Obj.SetDocument(NomeArq, CStr(HandleAnexo))
      Obj.Select("")
      Set Obj = Nothing
      Commit
      Set QWork = Nothing
    End If
  End If
End Sub

Public Sub TABLE_AfterInsert()
  Dim qFilial As BPesquisa
  Set qFilial = NewQuery
  qFilial.Add("SELECT FILIALPADRAO FROM Z_GRUPOUSUARIOS WHERE HANDLE= :PUSUARIO")
  qFilial.ParamByName("PUSUARIO").Value = CurrentUser
  qFilial.Active = True
  CurrentQuery.FieldByName("FILIAL_RESPONSAVEL").Value = qFilial.FieldByName("FILIALPADRAO").AsString

  Set qFilial = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim qLeituraInform As BPesquisa
  Set qLeituraInform = NewQuery

  qLeituraInform.Add("SELECT COUNT(1) QTDELIDOS FROM SAM_INFORMATIVO_USUARIOLEITURA WHERE INFORMATIVO = :HINFORMATIVO")
  qLeituraInform.ParamByName("HINFORMATIVO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qLeituraInform.Active = True

  If qLeituraInform.FieldByName("QTDELIDOS").AsInteger > 0 Then
  	CanContinue = False
	bsShowMessage("Informativo já lido não pode ser excluído.", "E")
  End If

  Set qLeituraInform = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (Not CurrentQuery.FieldByName("DATAFIM").IsNull) And _
      (CurrentQuery.FieldByName("DATAFIM").AsDateTime < CurrentQuery.FieldByName("DATAINICIO").AsDateTime) Then
    CanContinue = False
    bsShowMessage("A Data final, se informada, deve ser maior ou igual a inicial", "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_NewRecord()
  Dim ANO As Long
  Dim SEQUENCIA As Long
  ANO = Format(ServerDate, "yyyy")

  CurrentQuery.FieldByName("ANO").Value = Str(ANO) + "-01-01"
  NewCounter("SAM_INFORMATIVO", ANO, 1, SEQUENCIA)
  CurrentQuery.FieldByName("NUMERO").Value = SEQUENCIA

  If (VisibleMode And NodeInternalCode = 0) Or (WebMode And WebMenuCode = "T5707") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 0
  End If
  If (VisibleMode And NodeInternalCode = 1) Or (WebMode And WebMenuCode = "T2960") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 1
  End If
  If (VisibleMode And NodeInternalCode = 2) Or (WebMode And WebMenuCode = "T2951") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 2
  End If
  If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebMenuCode = "T2952") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 3
  End If
  If (VisibleMode And NodeInternalCode = 4) Or (WebMode And WebMenuCode = "T2953") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 4
  End If
  If (VisibleMode And NodeInternalCode = 5) Or (WebMode And WebMenuCode = "T4296") Then
	 CurrentQuery.FieldByName("TABCOBERTURA").Value = 5
  End If
  If (VisibleMode And NodeInternalCode = 6) Then
  	CurrentQuery.FieldByName("TABCOBERTURA").Value = NodeInternalCode
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOANEXARDOC") Then
		BOTAOANEXARDOC_OnClick
	End If
End Sub

'HASH: 6788CC8223993CCD9FB9FFA43B702150
'#Uses "*bsShowMessage"


Public Sub PAGAMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_CALENDGERAL_RECEBIMENTO")

  ShowPopup = False

  Dim datapag As Date

  datapag = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime


  PAGAMENTO.LocalWhere = " DATAFECHAMENTO IS NULL AND DATAPAGAMENTO >= " + SQLDate(datapag)

  ShowPopup = True
End Sub

Public Sub TABLE_AfterInsert()
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT DATAFINAL FROM SAM_CALENDGERAL_RECEBIMENTO WHERE CALENDGERAL = :CALENDGERAL ORDER BY DATAFINAL DESC")
  Q.ParamByName("CALENDGERAL").AsInteger = CurrentQuery.FieldByName("CALENDGERAL").AsInteger
  Q.Active = True
  If Not Q.FieldByName("DATAFINAL").IsNull Then
    CurrentQuery.FieldByName("DATAINICIAL").Value = (Q.FieldByName("DATAFINAL").AsDateTime + 1)
    DATAINICIAL.ReadOnly = True
    DATAFINAL.SetFocus
  Else
    DATAINICIAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  If VisibleMode Then
    Set Interface = CreateBennerObject("samcalendariopgto.ROTINAS")
    Interface.INICIALIZAR(CurrentSystem)

    'If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime<>Interface.DIAUTILANTERIOR(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)Then
    'MsgBox("Entre com um dia útil para a Data de Pagamento")
    'DATAPAGAMENTO.SetFocus
    'Interface.FINALIZAR
    'Set Interface=Nothing
    'CanContinue=False
    'Exit Sub
    'End If

    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
    Linha = Interface.Vigencia(CurrentSystem, "SAM_CALENDGERAL_RECEBIMENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CALENDGERAL", "")

    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(Linha, "E")
    End If
    Set Interface = Nothing
  End If

  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT DATAPAGAMENTO FROM SAM_PAGAMENTO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PAGAMENTO").AsInteger
  Q.Active = True
  If Q.FieldByName("DATAPAGAMENTO").AsDateTime <= CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    If VisibleMode Then
    	bsShowMessage("Data de pagamento nao pode ser menor ou igual a data final", "E")
   		CanContinue = False
  	End If
  End If

End Sub


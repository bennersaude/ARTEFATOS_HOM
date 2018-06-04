'HASH: E905B014E70904CB4B4210F8C69BC5B2

Option Explicit


'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"
'#Uses "*ProcuraTabelaUS"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  '  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim vHandleGrau As Long
  ShowPopup = False
  vHandleGrau = ProcuraGrau(GRAU.Text)
  If vHandleGrau <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandleGrau
  End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraTabelaUS(TABELAUS.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAUS").Value = vHandle
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
    Condicao = Condicao + " AND REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString
    Condicao = Condicao + " AND REDERESTRITAPRESTADOR = " + CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsString
  Else
    Condicao = "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
    Condicao = Condicao + " AND REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString
    Condicao = Condicao + " AND REDERESTRITAPRESTADOR = " + CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsString
  End If
  Linha = interface.Vigencia(CurrentSystem, "SAM_PRECOREDEREGIME_GRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "REDERESTRITA", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    MsgBox(Linha)
  End If
  Set interface = Nothing

End Sub


'HASH: AE502A214C58779DA971875F97C6B2E9
'Macro: SAM_GUIA_DEVOLUCAO
'#Uses "*ProcuraBeneficiarioAtivo"
'#Uses "*ProcuraPrestador"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraBeneficiarioAtivo(False, ServerDate, BENEFICIARIO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", PRESTADOR.Text) ' pelo CPF e executor
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If


End Sub

Public Sub TABDEVOLUCAO_OnChanging(AllowChange As Boolean)
  If CurrentQuery.State <> 3 Then
    AllowChange = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABDEVOLUCAO").AsInteger>1 Then
    Exit Sub
  End If
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT P.PRESTADOR CNPJCPF, P.FISICAJURIDICA, P.HANDLE,P.NOME, E.LOGRADOURO, E.NUMERO, E.COMPLEMENTO, E.BAIRRO, E.CEP")
  q1.Add("  FROM SAM_PRESTADOR P, ")
  q1.Add("       SAM_PRESTADOR_ENDERECO E ")
  q1.Add(" WHERE P.HANDLE = :PRESTADOR")
  q1.Add("   AND E.PRESTADOR = P.HANDLE")
  q1.Add("   AND E.CORRESPONDENCIA = 'S'")
  q1.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  q1.Active = True
  If q1.EOF Then
    MsgBox("Prestador sem endereço cadastrado")
    CanContinue = False
  End If

End Sub

Public Sub TABLE_NewRecord()
  Dim PRFILIAL As Long
  Dim PRFILIALPROCESSAMENTO As Long
  Dim PRMSG As String
  If BuscarFiliais(CurrentSystem, PRFILIAL, PRFILIALPROCESSAMENTO, PRMSG) Then
    MsgBox ("Erro na rotina de Busca de filiais.")
    Exit Sub
  End If

  CurrentQuery.FieldByName("filialprocessamento").Value = PRFILIAL


End Sub


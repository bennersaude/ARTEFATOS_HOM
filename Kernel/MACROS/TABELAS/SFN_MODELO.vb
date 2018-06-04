'HASH: 18AEB09F20973F2A4DBC77D923A9DEA6
'Macro: SFN_MODELO

'#Uses "*bsShowMessage

'alteração: 05/12/2001
'             por: Milton
'             SMS: 5325
'             Sub: BOTAOEXPORTAR_OnClick()

'Última alteração: 18/06/2002
'             por: Milton
'             SMS: 10013
'             Sub: BOTAOEXCLUIR_OnClick()


Option Explicit

Public Sub BOTAOCOPIAR_OnClick()

  Dim interface As Object
  Dim psMensagem As String

  Set interface = CreateBennerObject("rotarq.rotinas")
  interface.CopiarEstrutura(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "Cópia de " + CurrentQuery.FieldByName("DESCRICAO").AsString, psMensagem)

  Set interface = Nothing

  bsShowMessage(psMensagem, "I")

  If VisibleMode Then
    RefreshNodesWithTable "SFN_MODELO"
  End If
End Sub

Public Sub BOTAOEXCLUIR_OnClick()

  If Not InTransaction Then StartTransaction

  'Verifica se o modelo está sendo usado em algum Tipo de Documento ou Rotina Arquivo'
  Dim sql3 As Object
  Set sql3 = NewQuery
  sql3.Add("SELECT count(x.handle) QuantModelo FROM")
  sql3.Add("  (SELECT  HANDLE ")
  sql3.Add("    FROM SFN_TIPODOCUMENTO TP")
  sql3.Add("    WHERE TP.MODELODOCUMENTO = :HANDLE")
  sql3.Add("  UNION ")
  sql3.Add("   SELECT HANDLE")
  sql3.Add("    FROM SFN_ROTINAARQUIVO RA")
  sql3.Add("    WHERE RA.MODELO =  :HANDLE) X")

  sql3.ParamByName("HANDLE").AsInteger =  CurrentQuery.FieldByName("HANDLE").AsInteger

  sql3.Active = True
  If sql3.FieldByName("QUANTMODELO").AsInteger >= 1 Then
    bsShowMessage("O Modelo está sendo usado, e não pode ser excluído.", "I")

  Else
    If bsShowMessage("Deseja excluir o Modelo de Documento?", "Q") = vbYes Then

      Dim SQL1 As Object
      Dim SQL2 As Object
      Set SQL1 = NewQuery
      Set SQL2 = NewQuery
      SQL1.Add("SELECT HANDLE FROM SFN_MODELO_ESTRUTURA WHERE MODELO= :HANDLE")
      SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL1.Active = True
      While Not SQL1.EOF
        SQL2.Clear
        SQL2.Add("DELETE FROM SFN_MODELO_ESTRUTURA_CAMPO WHERE MODELOESTRUTURA= :HANDLE")
        SQL2.ParamByName("HANDLE").AsInteger = SQL1.FieldByName("HANDLE").AsInteger
        SQL2.ExecSQL
        SQL1.Next
      Wend
      SQL1.Active = False
      SQL1.Clear
      SQL1.Add("DELETE FROM SFN_MODELO_ESTRUTURA WHERE MODELO= :HANDLE")
      SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL1.ExecSQL

      SQL1.Active = False
      SQL1.Clear
      SQL1.Add("DELETE FROM SFN_MODELO WHERE HANDLE= :HANDLE")
      SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL1.ExecSQL

      Set SQL1 = Nothing
      Set SQL2 = Nothing
    End If

    If InTransaction Then Commit
  End If
  Set sql3 = Nothing

  If VisibleMode Then
    RefreshNodesWithTable "SFN_MODELO"
  End If

End Sub

Public Sub BOTAOEXPORTAR_OnClick()

  Dim interface As Object
  Set interface = CreateBennerObject("rotarq.rotinas")
  interface.ExportarModelo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing


End Sub
Public Sub MODELORELATORIO_OnBtnClick()
  Dim LocalizaDLL As Object
  Dim vlHandleXX As Long

  On Error GoTo Cancel
  Set LocalizaDLL = CreateBennerObject("Procura.Procurar")
  vlHandleXX = LocalizaDLL.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Descrição|Código", "CODIGO LIKE 'SFN%'", "Procura por Modelo Relatório", True, "")

  If vlHandleXX <>0 Then
    Dim SQL As Object
    Set SQL =NewQuery
    SQL.Add("SELECT CODIGO FROM R_RELATORIOS WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = vlHandleXX
    SQL.Active = True
    If CurrentQuery.State =1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("MODELORELATORIO").AsString = SQL.FieldByName("CODIGO").AsString
    Set SQL = Nothing
  End If
  Set LocalizaDLL = Nothing
  Exit Sub


  Cancel :
  bsShowMessage("Erro ao escolher um Modelo Relatório!", "E")
End Sub
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Pinheiro sms 32077
  If CurrentQuery.FieldByName("TABTIPOSEQUENCIALENVIO").AsInteger = 2 Then
    If CurrentQuery.FieldByName("SEQUENCIALENVIOINICIAL").AsInteger >= CurrentQuery.FieldByName("SEQUENCIALENVIOFINAL").AsInteger Then
      bsShowMessage("Valor inicial do intervalo não pode ser maior ou igual ao valor final do intervalo.", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("SEQUENCIALENVIO").AsInteger < CurrentQuery.FieldByName("SEQUENCIALENVIOINICIAL").AsInteger Then
      bsShowMessage("Sequencial de envio não pode ser inferior ao valor inicial do intervalo.", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("SEQUENCIALENVIO").AsInteger > CurrentQuery.FieldByName("SEQUENCIALENVIOFINAL").AsInteger Then
      bsShowMessage("Sequencial de envio não pode ser superior ao valor final do intervalo.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  'sms 57218
  If (CurrentQuery.FieldByName("TABTIPO").AsInteger = 1) And (CurrentQuery.FieldByName("TABREMESSARETORNO").AsInteger = 1) Then  'Crislei.Sorrilha SMS 114342
    If (CurrentQuery.FieldByName("REMESSAALTERACAOVENCIMENTO").IsNull) And (CurrentQuery.FieldByName("REMESSABAIXA").IsNull) And (CurrentQuery.FieldByName("REMESSACANCELA").IsNull) And (CurrentQuery.FieldByName("REMESSAENTRADA").IsNull) Then
      bsShowMessage("É necessário informar ao menos um código de remessa, baixa, cancelamento ou alteração de vencimento!", "E")
      CanContinue = False
      REMESSAENTRADA.SetFocus
      Exit Sub
    Else
      If (CurrentQuery.FieldByName("REMESSAALTERACAOVENCIMENTO").AsInteger = 0) And (CurrentQuery.FieldByName("REMESSABAIXA").AsInteger = 0) And (CurrentQuery.FieldByName("REMESSACANCELA").AsInteger = 0) And (CurrentQuery.FieldByName("REMESSAENTRADA").AsInteger = 0) Then
        bsShowMessage("É necessário informar ao menos um código de remessa, baixa, cancelamento ou alteração de vencimento maior que zero!", "E")
        CanContinue = False
        REMESSAENTRADA.SetFocus
        Exit Sub

      End If
    End If
  End If
  'fim sms 57218

  'Se o modelo não for de remessa os campos de código de remessa devem ser zerados
  If CurrentQuery.FieldByName("TABTIPOMODELO").AsInteger     <> 1 Or _
     CurrentQuery.FieldByName("TABTIPO").AsInteger           <> 1 Or _
     CurrentQuery.FieldByName("TABREMESSARETORNO").AsInteger <> 1 Then
    CurrentQuery.FieldByName("REMESSAENTRADA").AsInteger = 0
    CurrentQuery.FieldByName("REMESSABAIXA").AsInteger = 0
    CurrentQuery.FieldByName("REMESSACANCELA").AsInteger = 0
    CurrentQuery.FieldByName("REMESSAALTERACAOVENCIMENTO").AsInteger = 0
  End If
End Sub
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
		Case "BOTAOCOPIAR"
			BOTAOCOPIAR_OnClick
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
		Case "BOTAOEXPORTAR"
			BOTAOEXPORTAR_OnClick
	End Select
End Sub

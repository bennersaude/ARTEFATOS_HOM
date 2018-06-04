'HASH: 759093CF3CD8C8DF1071C8C2398AE277
'Macro: SFN_VISAOESTRUTURA

Option Explicit

Public Sub BOTAOGERACLASSES_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.GeraClasses(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

'GeraTabelaFilha(pSys: OleVariant; pTabelaOrigemDados, pTabelaPai, pTabelaFilha, pCampoLigacaoOrigemDados,
'                pCampoLigacaoPai, pSqlWhereEspecial, pSqlWhereEspecialIn, pLabelForm, pCamposGrid: WideString;
'  pHandlePai: Integer);

' interface.GeraTabelaFilha(CurrentSystem,"SFN_CLASSEGERENCIAL", "SFN_VISAOESTRUTURA", "SFN_VISAOESTRUTURA_CLASSE", "CLASSEGERENCIAL", _
'                                          "VISAOESTRUTURA", "AND SFN_CLASSEGERENCIAL.ULTIMONIVEL='S'", "", _
'                                          "Gera Classes Gerenciais", "HANDLE|ESTRUTURA|DESCRICAO", "Handle|Estrutura|Descrição", CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "S" Then
    BOTAOGERACLASSES.Enabled = True
  Else
    BOTAOGERACLASSES.Enabled = False
  End If
End Sub


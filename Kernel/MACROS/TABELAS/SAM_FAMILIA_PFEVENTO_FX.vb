'HASH: B5FC4C8003147A867AE0A5AC250186FB
'Macro: SAM_FAMILIA_PFEVENTO_FX
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  Dim qOrdem
  Set qOrdem = NewQuery
  qOrdem.Active = False
  qOrdem.Clear
  qOrdem.Add("SELECT MAX(ORDEM) ORDEM FROM SAM_FAMILIA_PFEVENTO_FX WHERE FAMILIATABPFEVENTO = :pHANDLEPFEVENTO")
  qOrdem.ParamByName("pHANDLEPFEVENTO").AsInteger = CurrentQuery.FieldByName("FAMILIATABPFEVENTO").AsInteger
  qOrdem.Active = True
  CurrentQuery.FieldByName("ORDEM").AsInteger = qOrdem.FieldByName("ORDEM").AsInteger + 1
  qOrdem.Active = False
  Set qOrdem = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim qTipoPF
  Set qTipoPF = NewQuery
  qTipoPF.Active = False
  qTipoPF.Clear
  qTipoPF.Add("SELECT TIPOPFVARIAVEL FROM SAM_FAMILIA_PFEVENTO WHERE HANDLE = :pPFEVENTO")
  qTipoPF.ParamByName("pPFEVENTO").AsInteger = RecordHandleOfTable("SAM_FAMILIA_PFEVENTO")
  qTipoPF.Active = True
  If (qTipoPF.FieldByName("TIPOPFVARIAVEL").AsString = "Q") Then ' Por quantidade
    VALORMAXIMO.ReadOnly = True
    QTDMAXIMA.ReadOnly = False
  Else
    VALORMAXIMO.ReadOnly = False
    QTDMAXIMA.ReadOnly = True
  End If
  qTipoPF.Active = False
  Set qTipoPF = Nothing

  Dim qPreNatal
  Set qPreNatal = NewQuery
  qPreNatal.Active = False
  qPreNatal.Clear
  qPreNatal.Add("SELECT CONTROLEPORPRENATAL FROM SAM_FAMILIA_PFEVENTO WHERE HANDLE = :pPFEVENTO")
  qPreNatal.ParamByName("pPFEVENTO").AsInteger = RecordHandleOfTable("SAM_FAMILIA_PFEVENTO")
  qPreNatal.Active = True

  If (qPreNatal.FieldByName("CONTROLEPORPRENATAL").AsString = "S") Then
    CODIGOPFPRENATAL.Visible = True
    VALORPFPRENATAL.Visible = True
  Else
    CODIGOPFPRENATAL.Visible = False
    VALORPFPRENATAL.Visible = False
  End If
  Set qPreNatal = Nothing

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim SQL
  Set SQL = NewQuery
  SQL.Add("SELECT DATAFINAL FROM SAM_FAMILIA_PFEVENTO WHERE HANDLE = :HFAMILIAPFEVENTO")
  SQL.ParamByName("HFAMILIAPFEVENTO").Value = RecordHandleOfTable("SAM_FAMILIA_PFEVENTO")
  SQL.Active = True
  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("PF finalizada não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_FAMILIA_PFEVENTO_FX")
  End If

End Sub

'**************************************************************************************************************
'************ Alteração Para não deixar deletar ordem inferior sem antes deletar a superior *******************
'**************************************************************************************************************

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Q2 As Object
  Set Q2 = NewQuery
  Q2.Add("SELECT HANDLE                                   ")
  Q2.Add("  FROM SAM_FAMILIA_PFEVENTO_FX                  ")
  Q2.Add(" WHERE FAMILIATABPFEVENTO = :FAMILIATABPFEVENTO ")
  Q2.Add("   AND ORDEM >  :ORDEM                          ")
  Q2.ParamByName("FAMILIATABPFEVENTO").AsInteger = CurrentQuery.FieldByName("FAMILIATABPFEVENTO").AsInteger
  Q2.ParamByName("ORDEM").AsInteger = CurrentQuery.FieldByName("ORDEM").AsInteger
  Q2.Active = True
  If Not Q2.EOF Then
    bsShowMessage("Existe uma ou mais ordens superiores a esta!", "E")
    CanContinue = False
  End If
  Q2.Active = False
  Set Q2 = Nothing
  '************************** Fim da ALteração ******************************************************************
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qBuscaOrdem
  Set qBuscaOrdem = NewQuery
  qBuscaOrdem.Active = False
  qBuscaOrdem.Clear
  qBuscaOrdem.Add("SELECT HANDLE FROM SAM_FAMILIA_PFEVENTO_FX WHERE ORDEM = :pORDEM AND FAMILIATABPFEVENTO = :pPFEVENTO")
  If (CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
    qBuscaOrdem.Add("AND HANDLE <> :pHANDLE")
  End If
  qBuscaOrdem.ParamByName("pORDEM").AsInteger = CurrentQuery.FieldByName("ORDEM").AsInteger
  qBuscaOrdem.ParamByName("pPFEVENTO").AsInteger = CurrentQuery.FieldByName("FAMILIATABPFEVENTO").AsInteger
  If (CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
    qBuscaOrdem.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
  qBuscaOrdem.Active = True
  If (Not qBuscaOrdem.EOF) Then
    CanContinue = False
    bsShowMessage("Já existe um registro com essa ordem.", "E")
    Set qBuscaOrdem = Nothing
    Exit Sub
  End If
  Set qBuscaOrdem = Nothing
  'Balani SMS 47582 20/07/2005
  Dim qAux As Object
  Set qAux = NewQuery
  qAux.Active = False
  qAux.Clear
  qAux.Add("SELECT HANDLE FROM SAM_FAMILIA_PFEVENTO_FX WHERE TABVALORPF <> :TABVALORPF AND FAMILIATABPFEVENTO = :FAMILIATABPFEVENTO")
  qAux.ParamByName("TABVALORPF").AsInteger = CurrentQuery.FieldByName("TABVALORPF").AsInteger
  qAux.ParamByName("FAMILIATABPFEVENTO").AsInteger = CurrentQuery.FieldByName("FAMILIATABPFEVENTO").AsInteger
  qAux.Active = True
  If Not qAux.FieldByName("HANDLE").IsNull Then
    CanContinue = False
    bsShowMessage("Não é permitido cadastrar faixas de participação financeira com tipos diferentes.", "E")
    Set qAux = Nothing
    CurrentQuery.FieldByName("CODIGOPF").Clear
    CurrentQuery.FieldByName("VALORPF").Clear
    Exit Sub
  End If
  Set qAux = Nothing
  'final SMS 47582

  Dim qPreNatal
  Set qPreNatal = NewQuery
  qPreNatal.Active = False
  qPreNatal.Clear
  qPreNatal.Add("SELECT CONTROLEPORPRENATAL FROM SAM_FAMILIA_PFEVENTO WHERE HANDLE = :pPFEVENTO")
  qPreNatal.ParamByName("pPFEVENTO").AsInteger = RecordHandleOfTable("SAM_FAMILIA_PFEVENTO")
  qPreNatal.Active = True

  If qPreNatal.FieldByName("CONTROLEPORPRENATAL").AsString = "S" Then
    If CurrentQuery.FieldByName("TABVALORPF").AsInteger = 1 Then
      If CurrentQuery.FieldByName("CODIGOPFPRENATAL").IsNull Then
        CanContinue = False
        bsShowMessage("O Código da pf para Pré-Natal deve ser informado.", "E")
        Set qPreNatal = Nothing
        Exit Sub
      End If
    Else
      If CurrentQuery.FieldByName("VALORPFPRENATAL").IsNull Then
        CanContinue = False
        bsShowMessage("O valor da pf para Pré-Natal deve ser informado.", "E")
        Set qPreNatal = Nothing
        Exit Sub
      End If
    End If
  End If

  Set qPreNatal = Nothing

End Sub



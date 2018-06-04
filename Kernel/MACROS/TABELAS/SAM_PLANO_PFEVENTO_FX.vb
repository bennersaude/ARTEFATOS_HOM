'HASH: 8C6419CE1EB0B1B19CEAD1EAE2414E08
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Q2 As Object
  '**************************************************************************************************************
  '************ Alteração Para não deixar deletar ordem inferior sem antes deletar a superior *******************
  '**************************************************************************************************************
  Set Q2 = NewQuery
  Q2.Add("SELECT HANDLE                         ")
  Q2.Add("  FROM SAM_PLANO_PFEVENTO_FX          ")
  Q2.Add(" WHERE PLANOPFEVENTO = :PLANOPFEVENTO ")
  Q2.Add("   AND ORDEM >  :ORDEM                ")
  Q2.ParamByName("PLANOPFEVENTO").AsInteger = CurrentQuery.FieldByName("PLANOPFEVENTO").AsInteger
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
  'SMS 64046 - ÉVERTON
  Dim qAux As Object

  Set qAux = NewQuery
  qAux.Active = False
  qAux.Clear
  qAux.Add("SELECT HANDLE FROM SAM_PLANO_PFEVENTO_FX WHERE TABVALORPF <> :TABVALORPF AND PLANOPFEVENTO = :PLANOPFEVENTO")
  qAux.ParamByName("TABVALORPF").AsInteger = CurrentQuery.FieldByName("TABVALORPF").AsInteger
  qAux.ParamByName("PLANOPFEVENTO").AsInteger = CurrentQuery.FieldByName("PLANOPFEVENTO").AsInteger
  qAux.Active = True

  If Not qAux.FieldByName("HANDLE").IsNull Then
    CanContinue = False
    bsShowMessage("Não é permitido cadastrar faixas de participação financeira com tipos diferentes.", "E")
    CurrentQuery.FieldByName("CODIGOPF").Clear
    CurrentQuery.FieldByName("VALORPF").Clear
  End If
  Set qAux = Nothing
  'FIM SMS 64046 - ÉVERTON
End Sub

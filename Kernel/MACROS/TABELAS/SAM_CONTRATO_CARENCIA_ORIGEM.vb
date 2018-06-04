'HASH: D9B6928FD6263BDD4918DF10928D87C0
'#Uses "*bsShowMessage"
'Daniela -02/08/2002

Public Sub ORIGEMCARENCIA_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim Handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  Handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_ORIGEMCARENCIA|SAM_ORIGEMCARENCIA[SAM_CONTRATO_ORIGEMCARENCIA.ORIGEMCARENCIA = SAM_ORIGEMCARENCIA.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Origem", True, "")
  If Handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ORIGEMCARENCIA").Value = Handlexx 'CONTRATOCARENCIA
  End If
  Set Procura = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim qOrigemCarencia As BPesquisa
Set qOrigemCarencia = NewQuery

qOrigemCarencia.Add("SELECT *                                     ")
qOrigemCarencia.Add("  FROM SAM_CONTRATO_CARENCIA_ORIGEM          ")
qOrigemCarencia.Add(" WHERE HANDLE <> :HANDLE                     ")
qOrigemCarencia.Add("   AND TIPODEPENDENTE = :TIPODEPENDENTE      ")
qOrigemCarencia.Add("   AND CONTRATOCARENCIA = :CONTRATOCARENCIA  ")

qOrigemCarencia.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
qOrigemCarencia.ParamByName("TIPODEPENDENTE").AsInteger = CurrentQuery.FieldByName("TIPODEPENDENTE").AsInteger
qOrigemCarencia.ParamByName("CONTRATOCARENCIA").AsInteger = CurrentQuery.FieldByName("CONTRATOCARENCIA").AsInteger
qOrigemCarencia.Active = True

If Not qOrigemCarencia.EOF Then
  bsShowMessage("Já foi inserido essa origem de carencia para esse tipo de dependente","E")
  CanContinue = False
  Exit Sub
End If
End Sub

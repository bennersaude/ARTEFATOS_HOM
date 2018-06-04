'HASH: 6408950A185F3ED0D50B6C139BBA86DB
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

Dim qOrigemCarencia As Object
Set qOrigemCarencia = NewQuery
qOrigemCarencia.Clear
qOrigemCarencia.Add("SELECT ORIGEMCARENCIA FROM SAM_FAMILIA_CARENCIA_ORIGEM WHERE ORIGEMCARENCIA = :ORIGEM AND FAMILIACARENCIA = :CARENCIA AND HANDLE <> :HANDLE")
qOrigemCarencia.ParamByName("CARENCIA").Value = CurrentQuery.FieldByName("FAMILIACARENCIA").AsInteger
qOrigemCarencia.ParamByName("ORIGEM").Value = CurrentQuery.FieldByName("ORIGEMCARENCIA").AsInteger
qOrigemCarencia.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
qOrigemCarencia.Active = True

If Not qOrigemCarencia.FieldByName("ORIGEMCARENCIA").IsNull Then
  bsShowMessage("Origem da Carência ja existente para esta Carência!!", "I")
  CanContinue = False
End If

Set qOrigemCarencia = Nothing

End Sub

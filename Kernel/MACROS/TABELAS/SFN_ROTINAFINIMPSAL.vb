'HASH: E8832FC36987375B137F41631459D6DB
'Macro: SFN_ROTINAFINIMPSAL


Public Sub BOTAOPROCESSAR_OnClick()
  Dim InterProc As Object
  Set InterProc = CreateBennerObject("SamImpSal.RotImpSal") 'sms 33156
  InterProc.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set InterProc = Nothing

  WriteAudit("P", HandleOfTable("SFN_ROTINAFINIMPSAL"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Importação de Salários - Processamento")
End Sub


Public Sub BOTAOCANCELAR_OnClick()
  Dim InterProc As Object
  Set InterProc = CreateBennerObject("SamImpSal.RotImpSal") 'sms 33156
  InterProc.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set InterProc = Nothing

  WriteAudit("C", HandleOfTable("SFN_ROTINAFINIMPSAL"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Importação de Salários - Cancelamento")
End Sub

Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  CONTRATOFINAL.LocalWhere = "SAM_CONTRATO.CONTRATO >= " + _
                             "(Select CONTRATO FROM SAM_CONTRATO WHERE SAM_CONTRATO.HANDLE = " + _
                             CurrentQuery.FieldByName("CONTRATOINICIAL").AsString + ")"
End Sub

Public Sub TABLE_AfterScroll()
  CONTRATOINICIAL.ResultFields = "CONTRATO|CONTRATANTE|"
  CONTRATOFINAL.ResultFields = "CONTRATO|CONTRATANTE|"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("CONTRATOINICIAL").AsInteger <>o And CurrentQuery.FieldByName("CONTRATOFINAL").AsInteger = 0 Then
    MsgBox("Você deve ter um Contrato Final quando tiver um  Contrato Inicial !   ")
    CanContinue = False
    Exit Sub
  End If
End Sub

'Inserindo no Campo SEQUENCIASALARIO a COMPETENCIA

Public Sub TABLE_NewRecord()
  Dim var_data As Date
  Dim var_String As String
  Dim sql As Object
  Set sql = NewQuery
  SQL.Add("SELECT COMPETENCIA FROM SFN_COMPETFIN WHERE HANDLE = :HANDLE")
  SQL.ParamByName("Handle").Value = RecordHandleOfTable("SFN_COMPETFIN")
  SQL.Active = True
  var_data = SQL.FieldByName("COMPETENCIA").AsString
  var_String = FormatDateTime2("YYYYMM", var_data) + "0"
  CurrentQuery.FieldByName("SEQUENCIASALARIO").Value = var_String
End Sub


'HASH: 51A94BFBBA199B42FADA5B3D6A0C6EA8
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
If CurrentQuery.State <> 1 Then
    bsShowMessage("Operação não permitida. O Registro não pode estar em edição", "I")
   	Exit Sub
  End If
  If Not (CurrentQuery.FieldByName("SITUACAO").AsString = "4")Then
    CanContinue = False
    bsShowMessage("Operação não permitida. A Rotina já foi processada.", "E")
    Exit Sub
  Else
   	CanContinue = True
  End If
  If BsShowMessage("Confirma o cancelamento da rotina?","Q") = vbYes Then
	If VisibleMode Then
	  Dim Dll As Object
	  Set Dll = CreateManagedObject("Benner.Saude.ProcContas.ExportaNFTS", "Benner.Saude.ProcContas.ExportaNFTS.CancelarRotinaExpNFTS")
	  Advert = Dll.Cancelar(CurrentSystem, _
	                        CurrentQuery.FieldByName("HANDLE").AsInteger)
	  Set Dll = Nothing

      If Advert Then
	    BSShowMessage("Verifique as ocorrências da rotina!", "I")
	  Else
		BSShowMessage("Rotina cancelada com sucesso!", "I")
	  End If
	End If
  End If
  RefreshNodesWithTable("SAM_ROTEXPORTANFTS")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
Dim Advert As Boolean
  If CurrentQuery.State <> 1 Then
    bsShowMessage("Operação não permitida. O Registro não pode estar em edição", "I")
    Exit Sub
  End If
  If Not (CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "2")Then
    CanContinue = False
    bsShowMessage("Operação não permitida. A Rotina já foi processada.", "E")
    Exit Sub
  Else
   	CanContinue = True
  End If
  If VisibleMode Then
    Dim Dll As Object
	Set Dll = CreateManagedObject("Benner.Saude.ProcContas.ExportaNFTS", "Benner.Saude.ProcContas.ExportaNFTS.ProcessarRotinaExpNFTS")

	Advert = Dll.Processar(CurrentSystem, _
		                         CurrentQuery.FieldByName("HANDLE").AsInteger)
   	Set Dll = Nothing
   	If Advert Then
	  BsShowMessage("Verifique as ocorrências da rotina!", "I")
    Else
	  BSShowMessage("Rotina processada com sucesso!", "I")
    End If
  End If
  RefreshNodesWithTable("SAM_ROTEXPORTANFTS")
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
	BOTAOPROCESSAR.Enabled = False
	BOTAOCANCELAR.Enabled = True
  ElseIf CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
	BOTAOCANCELAR.Enabled = False
	BOTAOPROCESSAR.Enabled = True
  Else
	BOTAOPROCESSAR.Enabled = False
	BOTAOCANCELAR.Enabled = False
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
 If CurrentQuery.State <> 1 Then
    bsShowMessage("Operação não permitida. O Registro não pode estar em edição", "I")
   	Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim COMPETENCIA As Object
  Set COMPETENCIA = NewQuery

  COMPETENCIA.Add("SELECT HANDLE ")
  COMPETENCIA.Add("FROM SAM_ROTEXPORTANFTS")
  COMPETENCIA.Add("WHERE COMPETENCIA = :COMPETENCIA")
  COMPETENCIA.Add("      AND HANDLE <> :HANDLE")
  COMPETENCIA.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  COMPETENCIA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  COMPETENCIA.Active = True

  If COMPETENCIA.FieldByName("HANDLE").AsInteger > 0 Then
	bsShowMessage("Competência já cadastrada.", "E")
	CanContinue = False
	Set COMPETENCIA = Nothing
  End If

  Set COMPETENCIA = Nothing
End Sub

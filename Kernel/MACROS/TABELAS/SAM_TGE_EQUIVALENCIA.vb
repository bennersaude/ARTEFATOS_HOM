'HASH: 551EEA5F194DC47AD6ADA165C056B77C
'macro tebela SAM_TGE_EQUIVALENTE
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  If (VisibleMode) Or (WebMode) Then
    CurrentQuery.FieldByName("TGEORIGEM").AsInteger = RecordHandleOfTable("SAM_TGE")
  End If
End Sub

Public Sub TABLE_AfterPost()
	bsShowMessage("Para replicação dos dados equivalentes, deve ser processada a verificação número 27.", "I")
End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    TGEEQUIVALENTE.WebLocalWhere = "ULTIMONIVEL = 'S' "
  End If
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If bsShowMessage("Deseja excluir a equivalencia e os dados replicados?", "Q") = vbYes Then
		Dim obj As Object
		Set obj = CreateBennerObject("SAMREPLICTGE.EventoEquivalente")
		obj.DeletaRegistrosEquivalentes(CurrentSystem, CurrentQuery.FieldByName("TGEEQUIVALENTE").AsInteger)
		Set obj = Nothing
	End If
End Sub

Public Sub TABLE_NewRecord()
  If (VisibleMode) Or (WebMode) Then
    CurrentQuery.FieldByName("TGEORIGEM").Value = RecordHandleOfTable("SAM_TGE")
  End If
End Sub

Public Sub TGEEQUIVALENTE_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TGEEQUIVALENTE").Value = vHandle
  End If
  Set interface = Nothing
  If (VisibleMode) Or (WebMode) Then
    CurrentQuery.FieldByName("TGEORIGEM").AsInteger = RecordHandleOfTable("SAM_TGE")
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vHandleEvento As Long
  Dim qCorrelato As Object

  If (VisibleMode) Or (WebMode) Then
    vHandleEvento = RecordHandleOfTable("SAM_TGE")
    CurrentQuery.FieldByName("TGEORIGEM").AsInteger = vHandleEvento
  End If

  If vHandleEvento = CurrentQuery.FieldByName("TGEEQUIVALENTE").AsInteger Then
    bsShowMessage("Evento equivalente igual ao evento origem!", "E")
    CanContinue = False
    Exit Sub
  End If

  Set qCorrelato = NewQuery
  qCorrelato.Clear
  qCorrelato.Add("SELECT COUNT(1) EVENTO                    ")
  qCorrelato.Add("  FROM SAM_TGE_EQUIVALENCIA 				")
  qCorrelato.Add(" WHERE TGEORIGEM = :EVENTO  				")
  qCorrelato.Add("   AND TGEEQUIVALENTE = :EVENTOEQUIVALENTE")
  qCorrelato.Add("   AND HANDLE <> :HANDLE                  ")
  qCorrelato.ParamByName("EVENTO").AsInteger = vHandleEvento
  qCorrelato.ParamByName("EVENTOEQUIVALENTE").AsInteger = CurrentQuery.FieldByName("TGEEQUIVALENTE").AsInteger
  qCorrelato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qCorrelato.Active = True

  If qCorrelato.FieldByName("EVENTO").AsInteger > 0 Then
    bsShowMessage("Evento equivalente já existe!", "E")
    CanContinue = False
    Exit Sub
  End If

  Set qCorrelato = Nothing
End Sub


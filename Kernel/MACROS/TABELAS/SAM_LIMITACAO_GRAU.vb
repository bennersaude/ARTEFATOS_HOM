'HASH: D29F083D1D481A1DF004E3BE215B5E51
Public Sub GRAU_OnPopup(ShowPopup As Boolean)
    Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  vCriterios = " NOT SAM_GRAU.HANDLE IN (SELECT SAM_LIMITACAO_GRAU.GRAU FROM SAM_LIMITACAO_GRAU WHERE SAM_LIMITACAO_GRAU.LIMITACAO = " + Str(RecordHandleOfTable("SAM_LIMITACAO")) + ")"
  vCampos = "Grau|Descrição|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterios, "Tabela de Graus", True,GRAU.Text)

  If vHandle > 0 Then
  	CurrentQuery.FieldByName("GRAU").AsInteger   = vHandle
  	CurrentQuery.FieldByName("CODIGO").AsInteger = vHandle
  End If

  Set interface = Nothing
  ShowPopup = False
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		GRAU.WebLocalWhere = " NOT A.HANDLE IN (SELECT SAM_LIMITACAO_GRAU.GRAU FROM SAM_LIMITACAO_GRAU WHERE SAM_LIMITACAO_GRAU.LIMITACAO = " + Str(RecordHandleOfTable("SAM_LIMITACAO")) + ")"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CurrentQuery.FieldByName("CODIGO").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
End Sub

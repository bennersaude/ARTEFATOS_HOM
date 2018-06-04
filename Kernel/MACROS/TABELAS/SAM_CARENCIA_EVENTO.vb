'HASH: 5F38806F3A9122484221937099726CDD
'Macro: SAM_CARENCIA_EVENTO
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim qSexo As Object
  Set qSexo = NewQuery
  ShowPopup = False

  'Pega o sexo da carencia
  qSexo.Clear
  qSexo.Add("SELECT SEXO             ")
  qSexo.Add("  FROM SAM_CARENCIA     ")
  qSexo.Add(" WHERE HANDLE = :HANDLE ")
  qSexo.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_CARENCIA")
  qSexo.Active = True


  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_CBHPM.DESCRICAO|SAM_TGE.DESCRICAOABREVIADA|SAM_TGE.NIVELAUTORIZACAO"

  vCriterio = "SAM_TGE.ULTIMONIVEL = 'S' "

  If qSexo.FieldByName("SEXO").AsString = "F" Then
    vCriterio = vCriterio + " AND SAM_TGE.SEXO IN ('F', 'A')"
  ElseIf qSexo.FieldByName("SEXO").AsString = "M" Then
    vCriterio = vCriterio + " AND SAM_TGE.SEXO IN ('M', 'A')"
  End If

  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM|Descrição abreviada TGE|Nível"

  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]", vColunas, 2, vCampos, vCriterio, _
            "Tabela Geral de Eventos", False, "", "CA011.ConsultaTge")

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

  Set Interface = Nothing
  qSexo.Active = False
  Set qSexo = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_CARENCIA_EVENTO_596" Then
			EVENTOESTRUTURA.ReadOnly = True
			CARENCIA.ReadOnly = True
		End If
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT HANDLE FROM SAM_CARENCIA_EVENTO WHERE CARENCIA = :CARENCIA AND EVENTO = :EVENTO AND HANDLE <> :HANDLE")
  qSel.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
  qSel.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.Active = True

  If Not qSel.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Evento já cadastrado.", "E")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing

End Sub

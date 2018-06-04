'HASH: A69D08FAE759D6C2044EE91E34682BDA
'Macro: SAM_LIMITACAO_EVENTO
Option Explicit
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"
'#Uses "*IsInt"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer
  Dim vHandle As Long

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_CBHPM.DESCRICAO|SAM_TGE.DESCRICAOABREVIADA|SAM_TGE.NIVELAUTORIZACAO"
  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM|Descrição abreviada TGE|Nível"

  vCriterio = "SAM_TGE.ULTIMONIVEL = 'S' " + _
              " AND NOT SAM_TGE.HANDLE IN (SELECT EVENTO " + _
              "                              FROM SAM_LIMITACAO_EVENTO " + _
              "                             WHERE LIMITACAO = " + Str(RecordHandleOfTable("SAM_LIMITACAO")) + ")"

  If InStr(EVENTO.LocateText, "1.") = 1 Then
    vOrdem = 2
  Else
    If IsInt(TiraAcento(EVENTO.LocateText,True)) Then
      If InStr(EVENTO.LocateText, ".") > 0 Then
         vOrdem = 1
      Else
         vOrdem = 3
      End If
    Else
      vOrdem = 3
    End If
  End If


  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]",vColunas,vOrdem ,vCampos,vCriterio, _
  "Tabela Geral de Eventos",False,EVENTO.LocateText,"CA011.ConsultaTge")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

  Set Interface = Nothing

End Sub


Public Sub TABLE_AfterScroll()

	If WebMode Then
		If WebMenuCode = "T1224" Then
			LIMITACAO.ReadOnly = True
		ElseIf WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
	End If

End Sub

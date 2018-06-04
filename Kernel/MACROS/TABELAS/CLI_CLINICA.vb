'HASH: A56B7B0A0358461A9010F8881F9EE087


Public Sub BOTAOAGENDA_OnClick()
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.Agendamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, 0, 0, 0)
  Set AGENDA = Nothing

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
  vCriterio = "FISICAJURIDICA = 2"
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  If WebMode Then
    PRESTADOR.WebLocalWhere = " @ALIAS.FISICAJURIDICA = 2"
  Else
    PRESTADOR.LocalWhere = " SAM_PRESTADOR.FISICAJURIDICA = 2"
  End If

End Sub

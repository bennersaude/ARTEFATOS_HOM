'HASH: FB3BB90ADAFA925EA177E67F006F8F00

'MACRO CLI_ROTALTERAESCALA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub MontaRotulo()
  ROTULODIASEMANA.Text = ""

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT DIASEMANA, TIPOATIVIDADE FROM CLI_ESCALA WHERE HANDLE = :ESCALA")
  SQL.ParamByName("ESCALA").Value = CurrentQuery.FieldByName("ESCALA").AsInteger
  SQL.Active = True

  If SQL.FieldByName("DIASEMANA").AsInteger = 1 Then
    ROTULODIASEMANA.Text = "Dia da semana: Domingo"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 2 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 3 Then
    ROTULODIASEMANA.Text = "Dia da semana: Terça"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 4 Then
    ROTULODIASEMANA.Text = "Dia da semana: Quarta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 5 Then
    ROTULODIASEMANA.Text = "Dia da semana: Quinta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 6 Then
    ROTULODIASEMANA.Text = "Dia da semana: Sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 7 Then
    ROTULODIASEMANA.Text = "Dia da semana: Sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 8 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 9 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 10 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a Domingo"
  End If

  If SQL.FieldByName("TIPOATIVIDADE").AsString = "A" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Ambos"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "P" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Procedimento"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "C" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Consulta"
  End If

  Set SQL = Nothing
End Sub


Public Sub BOTAOALTERAR_OnClick()
  Dim ClinicaDLL As Object
  Set ClinicaDLL = CreateBennerObject("CliClinica.Agenda")
  'Balani SMS 62412 18/05/2006
  If Not ((CurrentQuery.State = 2) Or (CurrentQuery.State = 3)) Then
    If CurrentQuery.FieldByName("DATAALTEROU").IsNull Then
      ClinicaDLL.AlteraEscala(CurrentSystem, CurrentQuery.FieldByName("ESCALA").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
      RefreshNodesWithTable("CLI_ROTALTERAESCALA")
    Else
      bsShowMessage("A rotina já foi executada uma vez!", "I")
    End If
  End If
  'final Balani SMS 62412 18/05/2006
  Set ClinicaDLL = Nothing
End Sub

Public Sub CLINICA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.SOLICITANTE|SAM_PRESTADOR.EXECUTOR|SAM_PRESTADOR.RECEBEDOR|SAM_PRESTADOR.NAOFATURARGUIAS"
  vCriterio = ""
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = interface.Exec(CurrentSystem, "CLI_CLINICA|SAM_PRESTADOR[SAM_PRESTADOR.HANDLE = CLI_CLINICA.PRESTADOR]", vColunas, 2, vCampos, vCriterio, "Clínica", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLINICA").Value = vHandle
    CurrentQuery.FieldByName("RECURSO").Clear
  End If
  Set interface = Nothing
End Sub



Public Sub RECURSO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  If CurrentQuery.FieldByName("CLINICA").AsInteger > 0 Then
    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.SOLICITANTE|SAM_PRESTADOR.EXECUTOR|SAM_PRESTADOR.RECEBEDOR|SAM_PRESTADOR.NAOFATURARGUIAS"
    vCriterio = "CLINICA = " + CurrentQuery.FieldByName("CLINICA").AsString
    vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
    vHandle = interface.Exec(CurrentSystem, "CLI_RECURSO|SAM_PRESTADOR[SAM_PRESTADOR.HANDLE = CLI_RECURSO.PRESTADOR]", vColunas, 2, vCampos, vCriterio, "Prestador", True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("RECURSO").Value = vHandle
    End If
    Set interface = Nothing
  End If
End Sub




Public Sub ESCALA_OnChange()
	MontaRotulo
End Sub



Public Sub TABLE_AfterScroll()
  MontaRotulo
  'Balani SMS 52412 18/05/2006
  If CurrentQuery.FieldByName("DATAALTEROU").IsNull Then
    BOTAOALTERAR.Visible = True
  Else
    BOTAOALTERAR.Visible = False
  End If
  'Balani SMS 52412 18/05/2006
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		ESCALA.WebLocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
	ElseIf VisibleMode Then
		ESCALA.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		ESCALA.WebLocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
	ElseIf VisibleMode Then
		ESCALA.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
	End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)



  If CurrentQuery.FieldByName("TABTIPOALTERACAO").AsInteger <> 3 Then
    Dim QESCALA As Object
    Set QESCALA = NewQuery

    QESCALA.Add("SELECT DATAINICIAL FROM CLI_ESCALA WHERE HANDLE = :ESCALA")
    QESCALA.ParamByName("ESCALA").Value = CurrentQuery.FieldByName("ESCALA").AsInteger
    QESCALA.Active = True

    If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <QESCALA.FieldByName("DATAINICIAL").AsDateTime Then
        bsShowMessage("A data final não pode ser anterior a data inicial!", "E")
        CanContinue = False
        Exit Sub
      End If

    End If

    If CurrentQuery.FieldByName("TABTIPOALTERACAO").AsInteger = 2 Then
      If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
        bsShowMessage("O campo data final é obrigatório!", "E")
        CanContinue = False
        Exit Sub
      Else
        If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < QESCALA.FieldByName("DATAINICIAL").AsDateTime Then
          bsShowMessage("A data final não pode ser anterior à data inicial da escala do recurso!", "E")
          Set QESCALA = Nothing
          CanContinue = False
          Exit Sub
        End If
      End If
    Else
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <= QESCALA.FieldByName("DATAINICIAL").AsDateTime Then
        bsShowMessage("A data inicial da nova escala não pode ser igual ou anterior da escala que está sendo alterada!", "E")
        CanContinue = False
        Exit Sub
      End If

      Dim Ok As Boolean
      Dim SQL As Object
      Set SQL = NewQuery
      SQL.Clear
      SQL.Add("SELECT CLINICA FROM CLI_RECURSO WHERE HANDLE = :RECURSO")
      SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
      SQL.Active = True
      Dim ClinicaDLL As Object
      Set ClinicaDLL = CreateBennerObject("CliClinica.Agenda")

      ClinicaDLL.EscalaSobreposta(CurrentSystem, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                                  CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
                                  CurrentQuery.FieldByName("HORAINICIALINTERVALO").AsDateTime, _
                                  CurrentQuery.FieldByName("HORAFINALINTERVALO").AsDateTime, _
                                  CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, _
                                  CurrentQuery.FieldByName("HORAFINAL").AsDateTime, _
                                  CurrentQuery.FieldByName("RECURSO").AsInteger, _
                                  CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
                                  CurrentQuery.FieldByName("ESCALA").AsInteger, _
                                  CurrentQuery.FieldByName("DIASEMANA").AsInteger, _
                                  SQL.FieldByName("CLINICA").AsInteger, _
                                  Ok)

      If Not Ok Then
        CanContinue = False
        Exit Sub
      End If
      Set ClinicaDLL = Nothing
    End If



    Set QESCALA = Nothing

  End If

  Dim VERIFICA As Object
  Set VERIFICA = NewQuery

  VERIFICA.Active = False
  VERIFICA.Clear
  VERIFICA.Add("SELECT 1 ")
  VERIFICA.Add("  FROM CLI_ROTALTERAESCALA ")
  VERIFICA.Add(" WHERE HANDLE <> :HANDLE ")
  VERIFICA.Add("   AND CLINICA = :CLINICA ")
  VERIFICA.Add("   AND RECURSO = :RECURSO ")
  VERIFICA.Add("   AND ESCALA = :ESCALA ")
  VERIFICA.Add("   AND USUARIOALTEROU IS NULL ")
  VERIFICA.Add("   AND DATAALTEROU IS NULL ")
  VERIFICA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  VERIFICA.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  VERIFICA.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  VERIFICA.ParamByName("ESCALA").AsInteger = CurrentQuery.FieldByName("ESCALA").AsInteger
  VERIFICA.Active = True

  If Not VERIFICA.EOF Then
    bsShowMessage("Já existe uma rotina de alteração para a escala selecionada!", "E")
    Set VERIFICA = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set VERIFICA = Nothing


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOALTERAR") Then
		BOTAOALTERAR_OnClick
	End If
End Sub

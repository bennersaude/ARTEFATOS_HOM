'HASH: DE40097F771F15EC4A0138E04470BAB3
'MACRO: MS_ATENDIMENTOAMBULATORIAL

Public Sub ATENDENTE_OnPopup(ShowPopup As Boolean)
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vHandle As Integer
  Dim Interface As Object
  Set Interface = CreateBennerObject("Procura.Procurar")
  ShowPopup = False

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.SOLICITANTE|SAM_PRESTADOR.EXECUTOR|SAM_PRESTADOR.RECEBEDOR|SAM_PRESTADOR.NAOFATURARGUIAS"
  vCriterio = "CLINICA="+CurrentQuery.FieldByName("CLINICA").AsString
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = Interface.Exec(CurrentSystem, "CLI_RECURSO|SAM_PRESTADOR[SAM_PRESTADOR.HANDLE=CLI_RECURSO.PRESTADOR]", vColunas, 2, vCampos, vCriterio, "Atendente", False, ATENDENTE.Text)

  If vHandle > 0 Then  CurrentQuery.FieldByName("ATENDENTE").AsInteger =  vHandle
End Sub

Public Sub PACIENTE_OnPopup(ShowPopup As Boolean)
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampo As String
  Dim vTabelas As String
  Dim vHandle As Integer
  Dim Interface As Object
  Set Interface = CreateBennerObject("Procura.Procurar")
  ShowPopup = False

  vColunas = "SAM_BENEFICIARIO.BENEFICIARIO|SAM_MATRICULA.NOME|SAM_MATRICULA.MATRICULA|MS_PACIENTES.IDADE"
  vCriterio = "MS_PACIENTES.FILIAL = " + Str(CurrentBranch)
  vCampos = "Beneficiário|Nome|Matrícula|Idade"
  vTabelas = "MS_PACIENTES|SAM_BENEFICIARIO[SAM_BENEFICIARIO.HANDLE=MS_PACIENTES.BENEFICIARIO]|SAM_MATRICULA[SAM_MATRICULA.HANDLE=MS_PACIENTES.MATRICULA]"

  vHandle = Interface.Exec(CurrentSystem, vTabelas, vColunas, 2, vCampos, vCriterio, "Paciente", False, PACIENTE.Text)

  If vHandle > 0 Then  CurrentQuery.FieldByName("PACIENTE").AsInteger =  vHandle
End Sub

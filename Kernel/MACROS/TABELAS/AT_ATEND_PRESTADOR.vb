'HASH: 1252F6D22EC9E090910F8336A2EC155E


Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim CLI As Object
  Dim HANDLECLI As Long
  Set CLI = NewQuery
  'PROCURA O HANDLE DO PRESTADOR QUE É A CLÍNICA
  CLI.Add("SELECT C.PRESTADOR                ")
  CLI.Add("  FROM AT_ATEND A,                ")
  CLI.Add("       AT_CLINICA C               ")
  CLI.Add(" WHERE A.CLINICA = C.HANDLE       ")
  CLI.Add("   AND A.HANDLE = :HANDLEATEND    ")
  CLI.ParamByName("HANDLEATEND").Value = CurrentQuery.FieldByName("ATENDIMENTO").Value
  CLI.Active = True
  HANDLECLI = CLI.FieldByName("PRESTADOR").AsInteger
  Set CLI = Nothing

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabelas As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  'Procura somente os prestadores que fazem parte do corpo clínico
  vTabelas = "SAM_PRESTADOR|SAM_PRESTADOR_PRESTADORDAENTID[SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PRESTADORDAENTID.PRESTADOR]"
  vColunas = "SAM_PRESTADOR.NOME|SAM_PRESTADOR.PRESTADOR" 'CAMPOS DA TABELA
  vCriterio = "ENTIDADE = " + Str(HANDLECLI)
  vCampos = "NOME|PRESTADOR" 'TÍTULO DOS CAMPOS

  vHandle = interface.Exec(CurrentSystem, vTabelas, vColunas, 1, vCampos, vCriterio, "Prestadores", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing

End Sub


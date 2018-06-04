'HASH: CBA6485B638A995B0F224FF4DC902154

'AT_PROGRAMA_PRESTADOR

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  If CurrentQuery.FieldByName("CLINICA").IsNull Then
    MsgBox("É necessário selecionar a clínica!")
    Exit Sub
  End If

  Dim CLI As Object
  Dim HANDLECLI As Long
  Set CLI = NewQuery
  'PROCURA O HANDLE DO PRESTADOR QUE É A CLÍNICA
  CLI.Add("SELECT PRESTADOR                ")
  CLI.Add("  FROM AT_CLINICA               ")
  CLI.Add(" WHERE HANDLE = :HANDLECLINICA    ")
  CLI.ParamByName("HANDLECLINICA").Value = CurrentQuery.FieldByName("CLINICA").Value
  CLI.Active = True
  HANDLECLI = CLI.FieldByName("PRESTADOR").AsInteger
  Set CLI = Nothing

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabelas As String

  Set interface = CreateBennerObject("Procura.Procurar")

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


'HASH: 1DC1540A93C3F457C0093D7D0BCA3B65

'Macro: SFN_CLASSEGERENCIAL_RECLASSALD

Public Sub CLASSEGERENCIALDESTINO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0 and ULTIMONIVEL = 'S'"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIALDESTINO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIALDESTINO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 2, vCampos, "", "Contratos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("CLASSEGERENCIAL").Value = RecordHandleOfTable("SFN_CLASSEGERENCIAL")
End Sub


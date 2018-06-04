'HASH: 169B353BA63A313847CDA74D20B1F371
'Macro: SAM_REDEDIFERENCIADA_PRESTADOR
'Mauricio Ibelli -06/02/2001 -sms1725 -Somente mostrar prestadores da rede diferenciada
'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  vCriterio = "SAM_PRESTADOR.REDEDIFERENCIADA = 'S'"
  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", True, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT REDEDIFERENCIADA FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.Active = True

  If SQL.FieldByName("REDEDIFERENCIADA").AsString = "N" Then
    bsShowMessage("Prestador não está cadastrado como rede diferenciada.", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT COUNT(*) PRESTADORESCADASTRADOS FROM SAM_REDEDIFERENCIADA_PRESTADOR     ")
  SQL.Add(" WHERE REDEDIFERENCIADA = :REDE AND PRESTADOR = :PRESTADOR AND HANDLE <> :HANDLE ")
  SQL.ParamByName("REDE").Value = CurrentQuery.FieldByName("REDEDIFERENCIADA").Value
  SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.Active = True

  If SQL.FieldByName("PRESTADORESCADASTRADOS").AsInteger > 0 Then
    bsShowMessage("Prestador já cadastrado para a rede diferenciada selecionada.", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing
End Sub


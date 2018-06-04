'HASH: 36FE0B3495842BB7E123E7DC9BB318E3

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim ProcuraDLL As Variant
  Dim vColunas As String
  Dim vCampos As String
  Dim vCriterio As String
  Dim vHandle As Long
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT FILIALAUDITOR FROM SAM_AUDITOR WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("AUDITOR").AsInteger
  sql.Active = True
  Set ProcuraDLL = CreateBennerObject("PROCURA.PROCURAR")
  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"
  vCriterio = "FILIALPADRAO = " + sql.FieldByName("FILIALAUDITOR").AsString
  vCampos = "CPF/CNPJ|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
  vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")
  ShowPopup = False
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  ShowPopup = False
  Set ProcuraDLL = Nothing
  Set sql = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 59477 - 22/03/2006 - Início
Dim Condicao As String
Dim Linha As String

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      MsgBox("A data inicial não pode ser posterior a data final!")
      CanContinue = False
      Exit Sub
    End If
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
  Condicao = Condicao + "AND AUDITOR = " + CurrentQuery.FieldByName("AUDITOR").AsString 'SMS 63895 - Marcelo Barbosa - 28/06/2006


  Linha = Interface.Vigencia(CurrentSystem, "SAM_AUDITOR_PRESTADOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)

    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      MsgBox(Linha)
    Exit Sub
  End If
  Set Interface = Nothing
'Rodrigo Soares - SMS: 59477 - 22/03/2006 - Fim
End Sub

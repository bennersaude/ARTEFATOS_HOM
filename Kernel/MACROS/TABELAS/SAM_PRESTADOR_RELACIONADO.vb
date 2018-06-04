'HASH: FDC6A17632CEE1A4ED932C154C3DB35D
'Macro: SAM_PRESTADOR_RELACIONADO
'#Uses "*bsShowMessage"

Dim vFiltro As String


Public Sub PRESTADORRELACIONADO_OnPopup(ShowPopup As Boolean)
  '  Dim interface As Object
  '  Dim vHandle As Long
  '  Dim vCampos As String
  '  Dim vColunas As String
  '  Dim vCriterio As String

  '  ShowPopup =False
  '  Set interface=CreateBennerObject("Procura.Procurar")
  '  vColunas ="SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  '  vColunas =vcolunas +"|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"
  '  vCriterio ="SAM_PRESTADOR.ASSOCIACAO='N'"

  '  vCampos ="CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

  '  vHandle =interface.Exec("SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]",vColunas,2,vCampos,vCriterio,"Prestadores",False,"")

  '  If vHandle<>0 Then
  '    CurrentQuery.Edit
  '    CurrentQuery.FieldByName("PRESTADORRELACIONADO").Value=vHandle
  '  End If
  '  Set INTERFACE=Nothing
  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If

  Dim ProcuraDLL As Variant
  Dim vColunas As String
  Dim vCampos As String
  Dim vCriterio As String
  Dim vHandle As Long
  Dim vUsuario As String
  vUsuario = Str(CurrentUser)
  vFilial = Str(CurrentBranch)
  Set ProcuraDLL = CreateBennerObject("PROCURA.PROCURAR")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vcolunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"
  vCriterio = "SAM_PRESTADOR.ASSOCIACAO='N' " + _
              "AND MUNICIPIOPAGAMENTO IN " + vFiltro

  vCampos = "CPF/CNPJ|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
  vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")
  ShowPopup = False
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADORRELACIONADO").Value = vHandle
  End If
  ShowPopup = False
  Set ProcuraDLL = Nothing

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface As Object
  Dim Linha As String
  Dim condicao As String

  condicao = " AND PRESTADORRELACIONADO = " + CurrentQuery.FieldByName("PRESTADORRELACIONADO").Value

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_RELACIONADO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set Interface = Nothing

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("Data Final nao pode ser menor que a data Inicial!", "E")
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "E", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub


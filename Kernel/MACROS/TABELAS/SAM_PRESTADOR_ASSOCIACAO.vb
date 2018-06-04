'HASH: 41A143187A59407CE208B968AAEF53B9
'Macro: SAM_PRESTADOR_ASSOCIACAO
Dim vFiltro As String
'#Uses "*bsShowMessage"


Public Sub ASSOCIACAO_OnPopup(ShowPopup As Boolean)

  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If


  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  If CurrentQuery.State = 2 Then
    vCriterio = "SAM_PRESTADOR.ASSOCIACAO = 'S' " + _
                "AND MUNICIPIOPAGAMENTO IN " + vFiltro
  End If
  If CurrentQuery.State = 3 Then
    vCriterio = "SAM_PRESTADOR.ASSOCIACAO = 'S' " + _
                "AND MUNICIPIOPAGAMENTO IN " + vFiltro
  End If

  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ASSOCIACAO").Value = vHandle
  End If
  ShowPopup = False
  Set interface = Nothing

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '  Dim interface As Object

  '  Dim vHandle As Long
  '  Dim vCampos As String
  '  Dim vColunas As String
  '  Dim vCriterio As String

  '  ShowPopup =False
  '  Set interface=CreateBennerObject("Procura.Procurar")

  '  vColunas ="SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  '  vColunas =vcolunas +"|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  '  vColunas ="SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  '  vColunas =vcolunas +"|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  '  vCriterio =""
  '  vCampos ="CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

  '  vHandle =interface.Exec("SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]",vColunas,2,vCampos,vCriterio,"Prestadores",True,"")

  '  If vHandle<>0 Then
  '    CurrentQuery.Edit
  '    CurrentQuery.FieldByName("PRESTADOR").Value=vHandle
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
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  'vCriterio ="SAM_PRESTADOR.ASSOCIACAO = 'S' " + _
  vCriterio = "SAM_PRESTADOR.ASSOCIACAO <> 'S' " + _
              "AND MUNICIPIOPAGAMENTO IN " + vFiltro

  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
  vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", True, "")
  ShowPopup = False
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  ShowPopup = False
  Set ProcuraDLL = Nothing
End Sub


Public Sub TABLE_AfterInsert()
  If (VisibleMode And _
      NodeInternalCode = 1) Or _
     (WebMode And _
      WebMenuCode = "T3874") Then 'Se for a visão de Associações
    CurrentQuery.FieldByName("ASSOCIACAO").Value = RecordHandleOfTable("SAM_PRESTADOR")
  Else
    CurrentQuery.FieldByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  End If
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  If WebMode Then
    If WebMenuCode = "T3874" Or WebMenuCode = "" Then 'Se for a visão de Associações
      ASSOCIACAO.ReadOnly = True
      ASSOCIACAO.Visible  = False
      PRESTADOR.ReadOnly  = False
      PRESTADOR.Visible   = True
    Else
      ASSOCIACAO.ReadOnly = False
      ASSOCIACAO.Visible  = True
      PRESTADOR.ReadOnly  = True
      PRESTADOR.Visible   = False
    End If
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qMunicipio As Object
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE = :HPRESTADOR")
  Q.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

  Q.Active = True
  If Q.FieldByName("ASSOCIACAO").AsString = "S" Then
    bsShowMessage("Este prestador não pode pertencer a uma associação, pois o mesmo é uma associação !", "E")
    CanContinue = False
    'RefreshNodesWithTable("SAM_PRESTADOR_ASSOCIACAO")
    Exit Sub
  End If
  Set Q = Nothing


  Set qMunicipio = NewQuery
  qMunicipio.Active = False
  qMunicipio.Clear
  qMunicipio.Add("SELECT MUNICIPIOPAGAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  qMunicipio.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  qMunicipio.Active = True
  If CurrentQuery.State = 2 Then
    If checkPermissao(CurrentSystem, CurrentUser, "M", qMunicipio.FieldByName("MUNICIPIOPAGAMENTO").AsInteger, "A") = "N" Then
      CanContinue = False
      bsShowMessage("Permissão negada! Usuário não pode alterar", "E")
      Set qMunicipio = Nothing
      Exit Sub
    End If
  End If
  If CurrentQuery.State = 3 Then
    If checkPermissao(CurrentSystem, CurrentUser, "M", qMunicipio.FieldByName("MUNICIPIOPAGAMENTO").AsInteger, "I") = "N" Then
      CanContinue = False
      bsShowMessage("Permissão negada! Usuário não pode incluir", "E")
      Set qMunicipio = Nothing
      Exit Sub
    End If
  End If
  Set qMunicipio = Nothing

  Dim interface As Object
  Dim Linha As String
  Dim CAMPO As String
  Dim CONDICAO As String

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  '********** Estava deixando cadastrar dois associados na mesma vigência Durval 14/11/2002 **************************
  'CONDICAO =" and classeassociado = " +"'" +CurrentQuery.FieldByName("classeassociado").AsString +"'"
  CONDICAO = " AND ASSOCIACAO = " + CurrentQuery.FieldByName("ASSOCIACAO").AsString
  '*******************************************************************************************************************

  Linha = interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ASSOCIACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", CONDICAO)

  If Linha = "" Then
    CONDICAO = ""
    Linha = interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ASSOCIACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", CONDICAO)
    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(Linha + " Para este Prestador em outra Associação", "E")
    End If
  Else
    CanContinue = False
    bsShowMessage(Linha + " Para este Prestador.", "E")
  End If
  Set interface = Nothing
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


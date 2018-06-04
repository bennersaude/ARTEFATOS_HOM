'HASH: 3487D970939109CA02481B679B31C050
'TV_FORM0143
'#Uses "*bsShowMessage"
'#Uses "*TV_FORM0143_VALIDACAO"

Option Explicit

Public Sub TABLE_AfterPost()
  If(WebMode)Then
    If Not(IniciarValidacoesTriagem(CurrentQuery.FieldByName("HANDLEPEG").AsString))Then
      Exit Sub
    End If
  End If
  RealizarProcessoEncaminhar
End Sub

Public Sub RealizarProcessoEncaminhar()
  PreencherEncaminhamento
  PreencherPeg
End Sub

Public Sub PreencherEncaminhamento()
  Dim sGuia As String
  Dim oQry As Object
  Set oQry = NewQuery
  oQry.Clear
  sGuia = BuscarModeloGuiaPorSetor(CurrentQuery.FieldByName("SETORORIGEM").Value)

  oQry.Add("INSERT INTO TRIAGEM_ENCAMINHAMENTO(HANDLE, PEG, SETOR, USUARIORESPONSAVEL, DATAENCAMINHAMENTO, SETORORIGEM, USUARIOORIGEM, REGRASETORMDGUIA, SETORDESTINO, USUARIODESTINO) ")
  oQry.Add(" VALUES(:HANDLE, :PEG, :SETOR, :USUARIORESPONSAVEL, :DATAENCAMINHAMENTO, :SETORORIGEM, :USUARIOORIGEM, :REGRASETORMDGUIA, :SETORDESTINO, :USUARIODESTINO)")
  oQry.ParamByName("HANDLE").Value                  = NewHandle("TRIAGEM_ENCAMINHAMENTO")
  oQry.ParamByName("PEG").Value                     = CurrentQuery.FieldByName("HANDLEPEG").Value
  oQry.ParamByName("SETOR").Value                   = CurrentQuery.FieldByName("SETORDESTINO").Value
  oQry.ParamByName("USUARIORESPONSAVEL").Value      = CurrentQuery.FieldByName("USUARIODESTINO").Value
  oQry.ParamByName("DATAENCAMINHAMENTO").AsDateTime = ServerNow
  oQry.ParamByName("SETORORIGEM").Value             = CurrentQuery.FieldByName("SETORORIGEM").Value
  oQry.ParamByName("USUARIOORIGEM").Value           = CurrentQuery.FieldByName("USUARIOORIGEM").Value
  oQry.ParamByName("REGRASETORMDGUIA").Value       = IIf(sGuia <> "", sGuia, Null)
  oQry.ParamByName("SETORDESTINO").Value            = CurrentQuery.FieldByName("SETORDESTINO").Value
  oQry.ParamByName("USUARIODESTINO").Value          = CurrentQuery.FieldByName("USUARIODESTINO").Value
  oQry.ExecSQL

  Set oQry = Nothing
End Sub

Public Function BuscarModeloGuiaPorSetor(handleSetor As String)
    Dim oQry As Object
    Dim Result As Boolean
    Set oQry = NewQuery
    oQry.Clear

    oQry.Add("SELECT MIN(HANDLE) AS MDGUIA FROM TRIAGEMSETOR_MDGUIA WHERE REGRASETOR IN (SELECT HANDLE FROM TRIAGEMSETOR_REGRA WHERE SETOR = :SETOR) ")
    oQry.ParamByName("SETOR").Value = handleSetor
    oQry.Active = True

    BuscarModeloGuiaPorSetor = oQry.FieldByName("MDGUIA").AsString
    Set oQry = Nothing
End Function

Public Sub PreencherPeg()
  Dim oQry As Object
  Set oQry = NewQuery
  oQry.Clear
  oQry.Add(" UPDATE SAM_PEG SET TRIAGEMSETOR = :TRIAGEMSETOR, TRIAGEMUSUARIO = :TRIAGEMUSUARIO WHERE HANDLE = :HANDLE")
  oQry.ParamByName("HANDLE").Value         = CurrentQuery.FieldByName("HANDLEPEG").Value
  oQry.ParamByName("TRIAGEMSETOR").Value   = CurrentQuery.FieldByName("SETORDESTINO").Value
  oQry.ParamByName("TRIAGEMUSUARIO").Value = CurrentQuery.FieldByName("USUARIODESTINO").Value
  oQry.ExecSQL

  Set oQry = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("HANDLEPEG").Value = CLng(SessionVar("HANDLE_PEG"))
  IniciarProcessos
End Sub

Public Sub IniciarProcessos()
  If(WebMode)Then
    IniciarValidacoesTriagem(CurrentQuery.FieldByName("HANDLEPEG").AsString)
  End If
  AlterarCamposDaEntidadePreenchido
  AlterarCamposDaEntidadeNaoPreenchido
End Sub

Public Sub AlterarCamposDaEntidadePreenchido()
  If Not(VerificarExisteEncaminhamento(CurrentQuery.FieldByName("HANDLEPEG").AsString)) Then
    Exit Sub
  End If
  BuscarUltimaTriagem
  PreencherFiltroSetorDestino
End Sub

Public Sub PreencherFiltroSetorDestino()
  Dim handle As String
  handle =  CurrentQuery.FieldByName("HANDLEPEG").AsString

  If(WebMode)Then
    SETORDESTINO.WebLocalWhere = RetornarFiltroSetorDestino(handle)
  ElseIf(VisibleMode)Then
    SETORDESTINO.LocalWhere = RetornarFiltroSetorDestino(handle)
  End If
End Sub

Public Sub BuscarUltimaTriagem()
    Dim oQry As Object
    Dim Result As Boolean
    Set oQry = NewQuery
    oQry.Clear

    oQry.Add("SELECT HANDLE, USUARIORESPONSAVEL, SETOR ")
    oQry.Add("FROM TRIAGEM_ENCAMINHAMENTO WHERE DATAENCAMINHAMENTO = (SELECT MAX(DATAENCAMINHAMENTO) FROM TRIAGEM_ENCAMINHAMENTO WHERE PEG = :PEG) AND PEG = :PEG")
    oQry.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLEPEG").AsInteger
    oQry.Active = True
    If(oQry.FieldByName("HANDLE").AsString = "")Then
      Set oQry = Nothing
      Exit Sub
    End If
    CurrentQuery.FieldByName("SETORORIGEM").Value = oQry.FieldByName("SETOR").AsString
    CurrentQuery.FieldByName("USUARIOORIGEM").Value = oQry.FieldByName("USUARIORESPONSAVEL").AsString
    SETORORIGEM.ReadOnly = True
    USUARIOORIGEM.ReadOnly = True
    Set oQry = Nothing
End Sub

Public Sub AlterarCamposDaEntidadeNaoPreenchido()
  If (VerificarExisteEncaminhamento(CurrentQuery.FieldByName("HANDLEPEG").AsString)) Then
    Exit Sub
  End If

  USUARIOORIGEM.ReadOnly = True

  PreencherFiltroSetorOrigem

  PreencherFiltroSetorDestino
End Sub

Public Sub PreencherFiltroSetorOrigem()
  Dim handle As String
  handle = CurrentQuery.FieldByName("HANDLEPEG").AsString

  If(WebMode)Then
    SETORORIGEM.WebLocalWhere = RetornarFiltroSetorOrigem(handle)
  ElseIf(VisibleMode)Then
    SETORORIGEM.LocalWhere = RetornarFiltroSetorOrigem(handle)
  End If

  If (CurrentQuery.FieldByName("USUARIOORIGEM").AsString = "") Then
    CurrentQuery.FieldByName("USUARIOORIGEM").Value = CurrentUser
  End If
End Sub

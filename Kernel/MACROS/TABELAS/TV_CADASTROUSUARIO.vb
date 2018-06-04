'HASH: 28E01B199263F988E4706A1D55E38A10
 
Option Explicit
Public Function UsuarioExiste(pCPF As String, pHandleUsuario As Long, pTipoUsuario As String) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.HANDLE, B.MATRICULAUNICA, C.PRESTADOR, D.PESSOA FROM Z_GRUPOUSUARIOS A ")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_BENEFICIARIO B ON (B.USUARIO = A.HANDLE)")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_PRESTADOR    C ON (C.USUARIO = A.HANDLE)")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_PESSOA       D ON (D.USUARIO = A.HANDLE)")
  SQL.Add("WHERE APELIDO = :APELIDO")
  SQL.ParamByName("APELIDO").AsString = pCPF
  SQL.Active = True
  If SQL.FieldByName("HANDLE").AsInteger > 0 Then
    UsuarioExiste = True
    pHandleUsuario = SQL.FieldByName("HANDLE").AsInteger
    If (SQL.FieldByName("MATRICULAUNICA").AsInteger > 0) And _
      (SQL.FieldByName("PRESTADOR").AsInteger > 0) Then
      pTipoUsuario = "A" ' Ambos
    End If
    If (SQL.FieldByName("MATRICULAUNICA").AsInteger > 0) Then
      pTipoUsuario = "B"
    End If

    If (SQL.FieldByName("PRESTADOR").AsInteger > 0) Then
      pTipoUsuario = "P"
    End If

    If (SQL.FieldByName("PESSOA").AsInteger > 0) Then
      pTipoUsuario = "E"
    End If

  Else
    UsuarioExiste = False
  End If

  Set SQL = Nothing
End Function



Public Function ValidarBeneficiario(pCPF As String, ByRef pNome As String, ByRef pEmail As String, ByRef pFilial As String, ByRef pMatricula As Long) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT M.NOME, M.CPF, B.EMAIL, FILIALCUSTO, M.HANDLE MATRICULA")
  SQL.Add("  FROM SAM_MATRICULA M")
  SQL.Add("  JOIN SAM_BENEFICIARIO B ON (B.MATRICULA = M.HANDLE)")
  SQL.Add(" WHERE M.CPF = :CPF")
  SQL.ParamByName("CPF").AsString = pCPF
  SQL.Active = True
  If Not SQL.EOF Then
    ValidarBeneficiario = True

    pNome = SQL.FieldByName("NOME").AsString
    pEmail = SQL.FieldByName("EMAIL").AsString
    pFilial = SQL.FieldByName("FILIALCUSTO").AsString
    pMatricula = SQL.FieldByName("MATRICULA").AsInteger

  Else
    ValidarBeneficiario = False
  End If
  Set SQL = Nothing
End Function

Public Function ValidarPrestador(pCPF As String, ByRef pNome As String, ByRef pEmail As String, ByRef pFilial As String, ByRef pHandlePrestador As Long) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT HANDLE,NOME, EMAIL, FILIALPADRAO")
  SQL.Add("  FROM SAM_PRESTADOR")
  SQL.Add(" WHERE CPFCNPJ = :CPF")
  SQL.ParamByName("CPF").AsString = pCPF
  SQL.Active = True
  If Not SQL.EOF Then
    ValidarPrestador = True
    pNome = SQL.FieldByName("NOME").AsString
    pEmail = SQL.FieldByName("EMAIL").AsString
    pFilial = SQL.FieldByName("FILIALPADRAO").AsString
    pHandlePrestador = SQL.FieldByName("HANDLE").AsInteger
  Else
    ValidarPrestador = False
  End If
  Set SQL = Nothing
End Function

Public Function ValidarPessoa(pCPF As String, ByRef pNome As String, ByRef pEmail As String, ByRef pFilial As Long, ByRef pHandlePessoa As Long) As Boolean
Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT HANDLE,NOME, EMAILRESPONSAVEL EMAIL")
  SQL.Add("  FROM SFN_PESSOA")
  SQL.Add(" WHERE CNPJCPF = :CPF")
  SQL.ParamByName("CPF").AsString = pCPF
  SQL.Active = True
  If Not SQL.EOF Then
    ValidarPessoa = True
    pNome = SQL.FieldByName("NOME").AsString
    pEmail = SQL.FieldByName("EMAIL").AsString
    pFilial = CurrentBranch
    pHandlePessoa = SQL.FieldByName("HANDLE").AsInteger
  Else
    ValidarPessoa = False
  End If
  Set SQL = Nothing

End Function


Public Function CriaUsuario(pNome As String, pGrupo As String, pEmail As String, pLogin As String, pFilial As String) As Long
  Dim vUser As CSSystemUser
  Dim viHandle As Long
  Set vUser = NewSystemUser
  vUser.NewUser
  vUser.UserProperty("NOME") = pNome
  vUser.UserProperty("GRUPO") = pGrupo
  vUser.UserProperty("EMAIL") = pEmail
  vUser.UserProperty("APELIDO") = pLogin
  vUser.UserProperty("SENHA") = pLogin
  vUser.UserProperty("FILIALPADRAO") = pFilial
  vUser.UserProperty("CODIGO") = "0"
  viHandle = vUser.SaveUserProperties
  vUser.SelectUser(viHandle)
  CriaUsuario = viHandle
  Set vUser = Nothing
End Function

Public Function LocalizaGrupo(pTipo As String) As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT GRUPOSEGURANCABENEFICIARIO, GRUPOSEGURANCAPRESTADOR FROM SAM_PARAMETROSWEB")
  SQL.Active = True

  If pTipo = "B" Then
    LocalizaGrupo = SQL.FieldByName("GRUPOSEGURANCABENEFICIARIO").AsInteger
  ElseIf pTipo = "P" Then
    LocalizaGrupo = SQL.FieldByName("GRUPOSEGURANCAPRESTADOR").AsInteger
  End If

  Set SQL = Nothing

End Function


Public Sub AdicionaGrupo(pGrupo As Long, pGrupoIncluido As Long, pUsuario As Long)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("INSERT INTO Z_GRUPOUSUARIOGRUPOS (HANDLE, GRUPO, GRUPOADICIONADO, USUARIO)")
  SQL.Add("VALUES (:HANDLE, :GRUPO, :GRUPOADICIONADO, :USUARIO)")
  SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOGRUPOS")
  SQL.ParamByName("GRUPO").AsInteger = pGrupo
  SQL.ParamByName("GRUPOADICIONADO").AsInteger = pGrupoIncluido
  SQL.ParamByName("USUARIO").AsInteger = pUsuario
  SQL.ExecSQL
  Set SQL = Nothing

End Sub

Public Sub VinculaBeneficiarioUsuario(pUsuario As Long, pMatriculaUnica As Long)
	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Clear
	SQL.Add("INSERT INTO Z_GRUPOUSUARIOS_BENEFICIARIO(HANDLE, USUARIO, MATRICULAUNICA)")
	SQL.Add("VALUES(:HANDLE, :USUARIO, :MATRICULA)")
	SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOS_BENEFICIARIO")
	SQL.ParamByName("USUARIO").AsInteger = pUsuario
	SQL.ParamByName("MATRICULA").AsInteger = pMatriculaUnica
	SQL.ExecSQL
End Sub

Public Sub VinculaPrestadorUsuario(pUsuario As Long, pPrestador As Long)
  Dim SQL As Object
	Set SQL = NewQuery
	SQL.Clear
	SQL.Add("INSERT INTO Z_GRUPOUSUARIOS_PRESTADOR(HANDLE, USUARIO, PRESTADOR)")
	SQL.Add("VALUES(:HANDLE, :USUARIO, :PRESTADOR)")
	SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOS_PRESTADOR")
	SQL.ParamByName("USUARIO").AsInteger = pUsuario
	SQL.ParamByName("PRESTADOR").AsInteger = pPrestador
	SQL.ExecSQL
End Sub


Public Sub TABLE_AfterPost()
  Dim vGrupoInserido As Long
  Dim vGrupo As Long
  Dim vNome As String
  Dim vEmail As String
  Dim vCPF As String
  Dim vFIlial As String
  Dim vUsuarioIncluido As Long
  Dim vMatricula  As Long
  Dim vHandlePrestador As Long
  Dim vTipoUsuario As String

  vCPF = CurrentQuery.FieldByName("CPFCNPJ").AsString
  vUsuarioIncluido = 0
  If Not UsuarioExiste(vCPF, vUsuarioIncluido, vTipoUsuario) Then

    If ValidarBeneficiario(vCPF, vNome, vEmail,vFIlial, vMatricula) Then
      vGrupo = LocalizaGrupo("B")
      vUsuarioIncluido = CriaUsuario( vNome, Str(vGrupo),vEmail,vCPF,vFIlial)
      VinculaBeneficiarioUsuario vUsuarioIncluido,vMatricula
      vGrupoInserido = vGrupo
    End If

    If ValidarPrestador(vCPF, vNome, vEmail, vFIlial, vHandlePrestador) Then
      vGrupo = LocalizaGrupo("P")
      If vUsuarioIncluido <= 0 Then
        vUsuarioIncluido = CriaUsuario(vNome, Str(vGrupo),vEmail,vCPF,vFIlial)
        VinculaPrestadorUsuario vUsuarioIncluido,vHandlePrestador
        vGrupoInserido = vGrupo
      Else
        AdicionaGrupo vGrupoInserido,vGrupo,vUsuarioIncluido
        VinculaPrestadorUsuario vUsuarioIncluido,vHandlePrestador
      End If
    End If
  Else
    If (vTipoUsuario = "P") Then
      If ValidarBeneficiario(vCPF, vNome, vEmail,vFIlial, vMatricula) Then
        VinculaBeneficiarioUsuario vUsuarioIncluido,vMatricula
        vGrupo = LocalizaGrupo("P")
        vGrupoInserido = LocalizaGrupo("B")
        AdicionaGrupo vGrupo,vGrupoInserido,vUsuarioIncluido
      End If
    End If

    If (vTipoUsuario = "B") Then
      If ValidarPrestador(vCPF, vNome, vEmail, vFIlial, vHandlePrestador) Then
        VinculaPrestadorUsuario vUsuarioIncluido,vHandlePrestador
        vGrupo = LocalizaGrupo("B")
        vGrupoInserido = LocalizaGrupo("P")
        AdicionaGrupo vGrupo,vGrupoInserido,vUsuarioIncluido
      End If
    End If

End If

End Sub

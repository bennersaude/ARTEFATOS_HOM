'HASH: 0BD398DE13F39071F71B728869F8C9A5
'Macro: SAM_PROPONENTE_EXPERIENCIA
'Mauricio Ibelli -14/08/2001 -sms3858 -acesso
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CEP_OnPopup(ShowPopup As Boolean)
  ' Joldemar Moreira 12/06/2003
  ' SMS 16059
  Dim vHandle As String
  Dim interface As Object
  ShowPopup = False
  Set interface = CreateBennerObject("ProcuraCEP.Rotinas")
  interface.Exec(CurrentSystem, vHandle)

  If vHandle <>"" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
    SQL.Add("  FROM LOGRADOUROS      ")
    SQL.Add(" WHERE CEP = :HANDLE ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
    CurrentQuery.FieldByName("LOGRADOUROCOMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("EXCLUIR").AsString <>"S" Then
   	bsShowMessage("Permissão negada! Usuário não pode excluir", "E")

      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("ALTERAR").AsString <>"S" Then
      bsShowMessage("Permissão negada! Usuário não pode alterar","E")

      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("INCLUIR").AsString <>"S" Then
      bsShowMessage("Permissão negada! Usuário não pode incluir", "E")

      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False

	  bsShowMessage("Data final anterior a data inicial", "E")
    End If
  End If
End Sub

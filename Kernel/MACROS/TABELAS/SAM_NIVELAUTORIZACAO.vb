'HASH: C50FFA0877D3F067B48C4032006D2D9F
Dim vghandle As Long
Dim vgNivel As Long

Public Sub TABLE_AfterPost()

  If VisibleMode = False Then 'Tratamento Integrator
    Exit Sub
  End If

  If CurrentQuery.FieldByName("NIVEL").AsInteger <> vgNivel Then
    Dim U As Object
    Set U = NewQuery
    U.Add("UPDATE SAM_TGE SET NIVELAUTORIZACAO = :N WHERE TABNIVELAUTORIZACAO = :H")
    U.ParamByName("H").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    U.ParamByName("N").Value = CurrentQuery.FieldByName("NIVEL").AsInteger
    U.ExecSQL
    U.Clear
    U.Add("UPDATE SAM_GRUPOUSUARIO SET NIVELAUTORIZACAO = :N WHERE NIVEL = :H")
    U.ParamByName("H").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    U.ParamByName("N").Value = CurrentQuery.FieldByName("NIVEL").AsInteger
    U.ExecSQL
    Set U = Nothing
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vghandle = CurrentQuery.FieldByName("HANDLE").AsInteger
  vgNivel = CurrentQuery.FieldByName("NIVEL").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Soares - SMS: 64141 - Ítem 4 - 29/06/2006 - Início
  Dim qVerifica As Object

  Set qVerifica = NewQuery
  qVerifica.Add("SELECT CODIGO              ")
  qVerifica.Add("  FROM SAM_NIVELAUTORIZACAO")
  qVerifica.Add(" WHERE CODIGO =  :CODIGO   ")
  qVerifica.Add("   AND HANDLE <> :HANDLE   ")
  qVerifica.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerifica.Active = True

  ' Verifica se já há alguem cadastrado com esse codigo, se tiver não aceita inserir novamente
  If (Not qVerifica.FieldByName("CODIGO").IsNull) Then
    MsgBox("Já existe um nivel de autorização com esse código")
    CanContinue = False
    CODIGO.SetFocus
    Set qVerifica = Nothing
    Exit Sub
  End If

  Set qVerifica = Nothing
  'Soares - SMS: 64141 - Ítem 4 - 29/06/2006 - Fim

End Sub

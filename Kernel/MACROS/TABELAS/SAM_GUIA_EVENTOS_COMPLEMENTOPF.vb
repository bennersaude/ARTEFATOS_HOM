'HASH: 0E6AF297AE2C843E292C4407C7AE770E
'#Uses "*bsShowMessage"
'#Uses "*VerificaPermissaoEdicaoTriagem"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  Dim sql2 As Object
  Set sql2 = NewQuery
  Dim v_VALORINFSUPPFOPERADORA As Double
  Dim v_VALORPF As Double


  If Not CurrentQuery.FieldByName("VALORINFSUPPFOPERADORA").IsNull Then

    v_VALORINFSUPPFOPERADORA = CurrentQuery.FieldByName("VALORINFSUPPFOPERADORA").Value

    sql.Active = False
    sql.Clear
    sql.Add("SELECT A.VALORPF                                                   ")
    sql.Add("  FROM SAM_GUIA_EVENTOS A, 										 ")
    sql.Add("       SAM_GUIA_EVENTOS_COMPLEMENTOPF B							 ")
    sql.Add(" WHERE A.HANDLE = B.GUIAEVENTO									 ")
    sql.Add("       And B.HANDLE = " + CurrentQuery.FieldByName("HANDLE").Value )
    sql.Active = True

    v_VALORPF = sql.FieldByName("VALORPF").AsInteger

    If v_VALORINFSUPPFOPERADORA > v_VALORPF Then
      MsgBox("O campo VALORINFSUPPFOPERADORA não pode ser maior que o VALORPF da guia !")
      CanContinue = False
      RefreshNodesWithTable("SAM_GUIA_EVENTOS_COMPLEMENTOPF")
      Set sql = Nothing
      Set sql2 = Nothing
      Exit Sub
    Else
      sql2.Active = False
      sql2.Clear
      sql2.Add(" UPDATE SAM_GUIA_EVENTOS_COMPLEMENTOPF                       ")
      sql2.Add("    SET VALORSUPPFOPERADORA    = :VALORNOVOSUPPF,            ")
      sql2.Add("        VALORINFSUPPFOPERADORA = :VALORNOVOSUPPF             ")
      sql2.Add("  WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").Value )
      sql2.ParamByName("VALORNOVOSUPPF").Value = v_VALORINFSUPPFOPERADORA
      sql2.ExecSQL
    End If
  End If

  Set sql = Nothing
  Set sql2 = Nothing

End Sub


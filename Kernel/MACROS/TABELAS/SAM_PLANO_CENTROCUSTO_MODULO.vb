'HASH: FB9B011EFD4D025C9F5AF04ADD123D5A
'MACRO SAM_PLANO_CENTROCUSTO_MODULO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If WebMode Then
  	MODULO.WebLocalWhere = "(HANDLE IN (SELECT MODULO FROM SAM_PLANO_MOD WHERE PLANO = (SELECT PLANO FROM SAM_PLANO_CENTROCUSTO WHERE HANDLE= @CAMPO(PLANOCENTROCUSTO))))"
  ElseIf VisibleMode Then
	MODULO.LocalWhere = "(HANDLE IN (SELECT MODULO FROM SAM_PLANO_MOD WHERE PLANO = (SELECT PLANO FROM SAM_PLANO_CENTROCUSTO WHERE HANDLE= @PLANOCENTROCUSTO)))"
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	MODULO.WebLocalWhere = "(HANDLE IN (SELECT MODULO FROM SAM_PLANO_MOD WHERE PLANO = (SELECT PLANO FROM SAM_PLANO_CENTROCUSTO WHERE HANDLE= @CAMPO(PLANOCENTROCUSTO))))"
  ElseIf VisibleMode Then
	MODULO.LocalWhere = "(HANDLE IN (SELECT MODULO FROM SAM_PLANO_MOD WHERE PLANO = (SELECT PLANO FROM SAM_PLANO_CENTROCUSTO WHERE HANDLE= @PLANOCENTROCUSTO)))"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT * FROM SAM_PLANO_CENTROCUSTO_MODULO")
  SQL.Add("WHERE HANDLE<>:HANDLE AND MODULO=:MODULO AND PLANOCENTROCUSTO=:HPLANO")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
  SQL.ParamByName("HPLANO").AsInteger = CurrentQuery.FieldByName("PLANOCENTROCUSTO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Este módulo já está cadastrado !", "E")
    CanContinue = False
  End If
  Set SQL = Nothing
End Sub


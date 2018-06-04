'HASH: E81CD8BA3E1E5CF9D17FC61EAB84DCA4
'Macro: SFN_ROTINAFINCOMISSAO
'#Uses "*bsShowMessage"


Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT SITUACAO")
  SQL.Add("FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :HROTINAFIN")
  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "A" Then
    bsShowMessage("A Rotina não foi processada", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 1 Then
    Set Obj = CreateBennerObject("SAMComissao.Apropriar")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")
  Else
    Set Obj = CreateBennerObject("SAMComissao.Faturar")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")
  End If
  Set Obj = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT SITUACAO")
  SQL.Add("FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :HROTINAFIN")
  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("A Rotina já foi processada", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  If CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 1 Then
    Set Obj = CreateBennerObject("SAMComissao.Apropriar")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "P")
  Else
    Set Obj = CreateBennerObject("SAMComissao.Faturar")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "P")
  End If
  Set Obj = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub

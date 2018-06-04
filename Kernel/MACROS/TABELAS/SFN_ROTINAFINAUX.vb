'HASH: 224BE712811A418146B6A689C1B5BF6A


'Macro da tabela SFN_ROTINAFINAUX
'Criada pela sms 55556 - Edilson.Castro - 28/01/2006

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  Set Obj = CreateBennerObject("SAMFaturamento.Faturamento")
  Obj.FaturarAuxilioAdiantamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set Obj = Nothing

  RefreshNodesWithTable("SFN_FATURA")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object

  Set Obj = CreateBennerObject("SAMFaturamento.Faturamento")
  Obj.FaturarAuxilioAdiantamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "P")
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set Obj = Nothing

  RefreshNodesWithTable("SFN_FATURA")
End Sub

Public Sub PEGFINAL_OnPopup(ShowPopup As Boolean)
  PEGFINAL.LocalWhere = "(SAM_PEG.PEG >= " + _
                        "(Select PEG FROM SAM_PEG WHERE SAM_PEG.HANDLE = " + _
                        Str(CurrentQuery.FieldByName("PEGINICIAL").AsInteger) + "))"  ' +  " AND " + _
                        '"(SAM_PEG.SITUACAO > '4' )"
End Sub

Public Sub PEGINICIAL_OnPopup(ShowPopup As Boolean)
  'PEGINICIAL.LocalWhere = "SAM_PEG.SITUACAO = ??? "
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAPAGTOINICIAL").AsDateTime >CurrentQuery.FieldByName("DATAPAGTOFINAL").AsDateTime Then
    CanContinue =False
    MsgBox("A data inicial não pode ser MAIOR do que a data final !")
    Exit Sub
  End If
End Sub
Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT SITUACAO            ")
  sql.Add("  FROM SFN_ROTINAFIN       ")
  sql.Add(" WHERE HANDLE = :HANDLE    ")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  sql.Active = True
  If sql.FieldByName("SITUACAO").AsString = "P" Then
    MsgBox("A Rotina já foi processada")
    Set sql = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set sql = Nothing
End Sub

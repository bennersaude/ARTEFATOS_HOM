'HASH: BBB883D12459E49DD7CF8E59304A77E7
'Macro: SFN_ROTINAFINIMPFAT
Option Explicit

Public Sub BOTAOPROCESSAR_OnClick()

  Dim Obj As Object
  Dim SQLRotFin As Object

  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True
  If SQLRotFin.FieldByName("SITUACAO").Value = "P" Then
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    MsgBox("A rotina já foi processada")
    Exit Sub
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing

  Set Obj = CreateBennerObject("SfnFatura.Rotinas")
  Obj.ProcessaImpFat(CurrentSystem, CurrentQuery.FieldByName("ROTINAFIN").AsInteger)
  Set Obj = Nothing

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  If MsgBox("Confirma o cancelamento da rotina ?", vbYesNo, "Rotina Financeira") = vbYes Then
    Dim Obj As Object
    Dim SQLRotFin As Object

    If CurrentQuery.State <>1 Then
      MsgBox("Os parâmetros não podem estar em edição")
      Exit Sub
    End If

    Set SQLRotFin = NewQuery
    SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
    SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
    SQLRotFin.Active = True
    If SQLRotFin.FieldByName("SITUACAO").Value <>"P" Then
      SQLRotFin.Active = False
      Set SQLRotFin = Nothing
      MsgBox("A rotina ainda não foi processada")
      Exit Sub
    End If
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing

    Set Obj = CreateBennerObject("Sfnfatura.rotinas")
    Obj.CancelaImpFat(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Obj = Nothing
  End If
End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim SQLRotFin As Object
  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True
  If SQLRotFin.FieldByName("SITUACAO").Value = "P" Then
    CanContinue = False
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    MsgBox("A Rotina já foi processada")
    Exit Sub
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Início
  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery
  With qRotinaFin
    .Active = False
    .Clear
    .Add("SELECT SITUACAO")
    .Add("  FROM SFN_ROTINAFIN")
    .Add(" WHERE HANDLE = :HANDLE")
    .ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    .Active = True
  End With
  If (qRotinaFin.FieldByName("SITUACAO").AsString = "P") Or (qRotinaFin.FieldByName("SITUACAO").AsString = "S") Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If
  Set qRotinaFin = Nothing
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Fim
End Sub

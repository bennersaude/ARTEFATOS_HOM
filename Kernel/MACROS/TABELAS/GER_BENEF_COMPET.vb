'HASH: 9EADED5494807098525526F5E77FDA67

Public Sub BOTAOGERAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição!")
    Exit Sub
  End If

  Dim Obj As Object

  Set Obj = CreateBennerObject("BSGER001.Rotinas")
  Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime, CurrentQuery.FieldByName("DATACALCULO").AsDateTime, _
           CurrentQuery.FieldByName("SOBREESCREVER").AsString, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Obj = Nothing
  Dim query As Object
  Set query = NewQuery
  query.Add("SELECT A.HANDLE FROM GER_BENEF_COMPETRESUMO A WHERE A.COMPETENCIA = :COMPETENCIA")
  query.ParamByName("COMPETENCIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.Active = True
  If query.EOF Then
    BOTAOGERAR.Enabled = True
    BOTAOCANCELAR.Enabled = False
  Else
    BOTAOGERAR.Enabled = False
    BOTAOCANCELAR.Enabled = True
  End If
  Set query = Nothing
End Sub


Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object


  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição!")
    Exit Sub
  End If


  Set Obj = CreateBennerObject("BSGER001.Rotinas")
  Obj.zera(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Dim query As Object
  Set query = NewQuery
  query.Add("SELECT A.HANDLE FROM GER_BENEF_COMPETRESUMO A WHERE A.COMPETENCIA = :COMPETENCIA")
  query.ParamByName("COMPETENCIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.Active = True
  If query.EOF Then
    BOTAOGERAR.Enabled = True
    BOTAOCANCELAR.Enabled = False
  Else
    BOTAOGERAR.Enabled = False
    BOTAOCANCELAR.Enabled = True
  End If
  Set query = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  Dim query As Object
  Set query = NewQuery
  query.Add("SELECT A.HANDLE FROM GER_BENEF_COMPETRESUMO A WHERE A.COMPETENCIA = :COMPETENCIA")
  query.ParamByName("COMPETENCIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.Active = True
  If query.EOF Then
    BOTAOGERAR.Enabled = True
    BOTAOCANCELAR.Enabled = False
  Else
    BOTAOGERAR.Enabled = False
    BOTAOCANCELAR.Enabled = True
  End If
  Set query = Nothing
End Sub

Public Sub BOTAOPLANILHA_OnClick()
  Dim Obj As Object
  Dim vCompetenciaInicial As Date
  Dim vCompetenciaFinal As Date


  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If


  Set Obj = CreateBennerObject("SamAnaliseCadastral.Rotinas")
  Obj.planilha(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

End Sub


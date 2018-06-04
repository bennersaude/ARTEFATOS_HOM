'HASH: 4421BD368A532C22E2368193CBF5194C
'Macro da tabela GER_MODELO_COMPET


Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Dim vCompetenciaInicial As Date
  Dim vCompetenciaFinal As Date


  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  vCompetenciaInicial = CurrentQuery.FieldByName("CompetenciaInicial").Value
  vCompetenciaFinal = CurrentQuery.FieldByName("CompetenciaFinal").Value

  Set Obj = CreateBennerObject("BSGER002.Rotinas")
  Obj.Totaliza(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Totalizacao = Nothing
  BOTAOGERAR.Enabled = False
  BOTAOCANCELAR.Enabled = True
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object
  Dim vCompetenciaInicial As Date
  Dim vCompetenciaFinal As Date


  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  vCompetenciaInicial = CurrentQuery.FieldByName("CompetenciaInicial").Value
  vCompetenciaFinal = CurrentQuery.FieldByName("CompetenciaFinal").Value

Dim Qtmp As Object
Set Qtmp = NewQuery
Qtmp.Active = False
Qtmp.Clear
Qtmp.Add("DELETE ")
Qtmp.Add("FROM GER_MODELO_COMPETRESUMO_ESPEC ")
Qtmp.Add("WHERE COMPETRESUMO In (Select HANDLE ")
Qtmp.Add("                          FROM GER_MODELO_COMPETRESUMO ")
Qtmp.Add("                         WHERE MODELOCOMPET = :HND) ")
Qtmp.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
Qtmp.ExecSQL
Set Qtmp = Nothing

  Set Obj = CreateBennerObject("ResumoBenef.Rotinas")
  Obj.limpa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  BOTAOCANCELAR.Enabled = False
  BOTAOGERAR.Enabled = True
End Sub

Public Sub TABLE_AfterScroll()
  Dim query As Object
  Set query = NewQuery
  query.Add("SELECT A.HANDLE FROM GER_MODELO_COMPETRESUMO A, GER_MODELO_COMPET B WHERE A.MODELOCOMPET = :MODELOCOMPET")
  query.ParamByName("MODELOCOMPET").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
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


  Set Obj = CreateBennerObject("SamAnaliseGerencial.Rotinas")
  Obj.planilha(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  If(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
     (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)Then
  MsgBox("A Competência Final , se informada, deve ser maior ou igual a inicial")
  CanContinue = False
Else
  CanContinue = True
End If


End Sub


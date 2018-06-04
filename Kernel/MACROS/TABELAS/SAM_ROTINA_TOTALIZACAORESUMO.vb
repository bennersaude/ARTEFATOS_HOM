'HASH: C12B324E998ED2183347F47C480D5C2E
'Macro: SAM_ROTINA_TOTALIZACAORESUMO
'Shiba 16/11/2001


Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Dim vCompetenciaInicial As Date
  Dim vCompetenciaFinal As Date
  Dim vHandleIdade As Integer
  Dim vHandle As Integer

  If CurrentQuery.State <>1 Then
    MsgBox("Os parÔmetros nÒo podem estar em ediþÒo")
    Exit Sub
  End If

  vCompetenciaInicial = CurrentQuery.FieldByName("CompetenciaInicial").Value
  vCompetenciaFinal = CurrentQuery.FieldByName("CompetenciaFinal").Value
  vHandleIdade = CurrentQuery.FieldByName("FAIXAS").Value
  vHandle = CurrentQuery.FieldByName("MODELO").Value

  Set Obj = CreateBennerObject("ResumoBenef.Rotinas")
  Obj.Totaliza(CurrentSystem, vCompetenciaInicial, vCompetenciaFinal, vHandleIdade, vHandle)
  Set Totalizacao = Nothing
  BOTAOGERAR.Enabled = False

End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'TOT001'")
  SQL.Active = False
  SQL.Active = True

  ReportPreview(SQL.FieldByName("HANDLE").AsInteger, "", False, False)
  Set SQL = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim QCOMISS As Object
  Dim QDESPESA As Object

  Set QCOMISS = NewQuery
  QCOMISS.Add("SELECT * FROM SAM_TOTALIZACAO_COMISS_MENSAL WHERE TOTALIZARESUMO = :HTOTALIZA")
  QCOMISS.ParamByName("HTOTALIZA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QCOMISS.Active = True

  Set QDESPESA = NewQuery
  QDESPESA.Add("SELECT * FROM SAM_TOTALIZACAODESPESA WHERE TOTALIZARESUMO = :HTOTALIZA")
  QDESPESA.ParamByName("HTOTALIZA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QDESPESA.Active = True

  If(Not QCOMISS.FieldByName("HANDLE").IsNull)And(Not QDESPESA.FieldByName("HANDLE").IsNull)Then
  BOTAOGERAR.Enabled = False
Else
  BOTAOGERAR.Enabled = True
End If

Set SQL = Nothing
Set SQL2 = Nothing

End Sub


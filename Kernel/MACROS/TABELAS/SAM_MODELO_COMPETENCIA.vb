'HASH: 419FA9BA30451C149973ADEBC122AE22
'Macro: SAM_MODELO_COMPETENCIA
'Origem: SAM_ROTINA_TOTALIZACAORESUMO -Shiba

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


  Set Obj = CreateBennerObject("ResumoBenef.Rotinas")
  Obj.Totaliza(CurrentSystem)
  Set Totalizacao = Nothing
  BOTAOGERAR.Enabled = False

End Sub

Public Sub TABLE_AfterScroll()
  Dim QCOMISS As Object
  Dim QTotalizaBenef As Object
  Dim QTotalizaBenefAberto As Object

  Set QCOMISS = NewQuery
  QCOMISS.Add("SELECT * FROM SAM_TOTALIZACAO_COMISS_MENSAL WHERE TOTALIZARESUMO = :HTOTALIZA")
  QCOMISS.ParamByName("HTOTALIZA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QCOMISS.Active = True

  Set QTotalizaBenef = NewQuery
  QTotalizaBenef.Add("SELECT * FROM SAM_TOT_BENEF WHERE COMPET = :HTOTALIZA")
  QTotalizaBenef.ParamByName("HTOTALIZA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QTotalizaBenef.Active = True

  Set QTotalizaBenefAberto = NewQuery
  QTotalizaBenefAberto.Add("SELECT * FROM SAM_TOT_BENEF_ABERTO WHERE TOTALBENEF = :HTOTALIZA")
  QTotalizaBenefAberto.ParamByName("HTOTALIZA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QTotalizaBenefAberto.Active = True

  If(Not QCOMISS.FieldByName("HANDLE").IsNull)And(Not QTotalizaBenef.FieldByName("HANDLE").IsNull)And(Not QTotalizaBenefAberto.FieldByName("HANDLE").IsNull)Then
  BOTAOGERAR.Enabled = False
Else
  BOTAOGERAR.Enabled = True
End If

Set SQL = Nothing
Set SQL2 = Nothing
Set SQL3 = Nothing

End Sub


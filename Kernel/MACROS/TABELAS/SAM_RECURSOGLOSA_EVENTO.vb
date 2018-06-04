'HASH: 33221A18048159C79CCA09D2489670A8
 

'MACRO: SAM_RECURSOGLOSA_EVENTO

'SMS 76000 - Marcelo Barbosa - 29/01/2007
Public Sub BOTAODEFERIR_OnClick()
  Dim CA001 As Object
  Set CA001 = CreateBennerObject("CA001.CAtend")

  CA001.DeferirRecursoGlosa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsFloat)

  Set CA001 = Nothing

  RefreshNodesWithTable("SAM_RECURSOGLOSA_EVENTO")
End Sub

Public Sub BOTAOINDEFERIR_OnClick()
  Dim CA001 As Object
  Set CA001 = CreateBennerObject("CA001.CAtend")

  CA001.IndeferirRecursoGlosa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsFloat)

  Set CA001 = Nothing
  RefreshNodesWithTable("SAM_RECURSOGLOSA_EVENTO")
End Sub

Public Sub TABLE_AfterScroll()
  Dim GlosaRecurso As Object
  Dim ParamProcContas As Object

  Set GlosaRecurso = NewQuery

  GlosaRecurso.Active = False
  GlosaRecurso.Clear
  GlosaRecurso.Add("SELECT SITUACAO FROM SAM_RECURSOGLOSA WHERE HANDLE = :HRECURSO")
  GlosaRecurso.ParamByName("HRECURSO").AsInteger = CurrentQuery.FieldByName("RECURSOGLOSA").AsInteger
  GlosaRecurso.Active = True

  If GlosaRecurso.FieldByName("SITUACAO").AsString = "P" Then
    Set ParamProcContas = NewQuery
    ParamProcContas.Active = False
    ParamProcContas.Clear
    ParamProcContas.Add("SELECT UTILIZADEFERIMENTO FROM SAM_PARAMETROSPROCCONTAS")
    ParamProcContas.Active = True

    If ParamProcContas.FieldByName("UTILIZADEFERIMENTO").AsString = "S" Then
      TIPORECURSO.ReadOnly = False
      If (CurrentQuery.FieldByName("MOTIVODEFERIMENTO").IsNull) And (CurrentQuery.FieldByName("MOTIVOINDEFERIMENTO").IsNull) Then
        BOTAODEFERIR.Enabled = True
        BOTAOINDEFERIR.Enabled = True
      Else
        BOTAODEFERIR.Enabled = False
        BOTAOINDEFERIR.Enabled = False
      End If
    Else
      TIPORECURSO.ReadOnly = True
    End If

    Set ParamProcContas = Nothing
  Else
    TIPORECURSO.ReadOnly = True
  End If

  Set GlosaRecurso = Nothing

  'Balani SMS 58046 21/12/2006
  If (CurrentQuery.FieldByName("DATADEFERIMENTO").IsNull) And (CurrentQuery.FieldByName("DATAINDEFERIMENTO").IsNull) Then
    TIPORECURSO.ReadOnly = False
  Else
    TIPORECURSO.ReadOnly = True
  End If
  'final Balani SMS 58046 21/12/2006
End Sub
'FIM SMS 76000

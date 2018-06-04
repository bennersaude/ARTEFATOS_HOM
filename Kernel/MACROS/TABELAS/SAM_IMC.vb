'HASH: 7F13B9D1B27F5DDDDA9A8DE5A972592B

' TABELA SAM_IMC
' COELHO - SMS 33275 - NOVEMBRO 2004 - CABESP
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qBusca As Object
  Dim vMinInf As Double
  Dim vMaxInf As Double
  Set qBusca = NewQuery

  vMinInf = CurrentQuery.FieldByName("VALORMINIMO").Value
  vMaxInf = CurrentQuery.FieldByName("VALORMAXIMO").Value

  If vMinInf >= vMaxInf Then
    bsShowMessage("Valor máximo deve ser maior que o valor mínimo!",  "E")
    VALORMAXIMO.SetFocus
    Set qBusca = Nothing
    CanContinue = False
    Exit Sub
  End If

  If vMinInf = 0 Then
    bsShowMessage("Valor mínimo deve ser informado!", "E" )
    VALORMINIMO.SetFocus
    qBusca.Active = False
    Set qBusca = Nothing
    CanContinue = False
    Exit Sub
  End If

  If vMaxInf = 0 Then
    bsShowMessage("Valor máximo deve ser informado!", "E")
    VALORMAXIMO.SetFocus
    qBusca.Active = False
    Set qBusca = Nothing
    CanContinue = False
    Exit Sub
  End If

  qBusca.Clear
  qBusca.Add("SELECT VALORMINIMO, VALORMAXIMO   ")
  qBusca.Add("  FROM SAM_IMC                    ")
  qBusca.Add(" WHERE HANDLE <>:HANDLE           ")
  qBusca.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qBusca.Active = True

  While Not qBusca.EOF
    If ((vMinInf >= qBusca.FieldByName("VALORMINIMO").Value) And (vMinInf <= qBusca.FieldByName("VALORMAXIMO").Value)) Then
      bsShowMessage("Valor mínimo já informado em outra faixa!", "E")
      VALORMINIMO.SetFocus
      qBusca.Active = False
      Set qBusca = Nothing
      CanContinue = False
      Exit Sub
    End If

    If ((vMinInf <= qBusca.FieldByName("VALORMINIMO").Value) And (vMaxInf >= qBusca.FieldByName("VALORMINIMO").Value)) Then
      bsShowMessage("Valor máximo ultrapassou o menor valor máximo já informado em outra faixa!", "E")
      VALORMAXIMO.SetFocus
      qBusca.Active = False
      Set qBusca = Nothing
      CanContinue = False
      Exit Sub
    End If

    qBusca.Next
  Wend

  Set qBusca = Nothing
End Sub


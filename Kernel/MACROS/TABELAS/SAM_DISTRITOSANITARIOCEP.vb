'HASH: A7542FFB941BFD4E24D0DA3E23749268
'Macro: SAM_DISTRITOSANITARIOCEP
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  Dim QueryVerificaCEP As Object
  Set QueryVerificaCEP = NewQuery

  'se intervalo correto
  If CurrentQuery.FieldByName("CEPINICIAL").Value > CurrentQuery.FieldByName("CEPFINAL").Value Then
    bsShowMessage("CEP Final <= Inicial", "E")
    CanContinue = False
    Exit Sub
  Else
    CanContinue = True
  End If

  'se cepinicial existe
  QueryVerificaCEP.Clear
  QueryVerificaCEP.Add("SELECT CEP FROM LOGRADOUROS")
  QueryVerificaCEP.Add("WHERE  CEP = :CEPINICIAL")
  QueryVerificaCEP.Add("  AND  HANDLE <> :HANDLE")
  QueryVerificaCEP.ParamByName("CEPINICIAL").Value = CurrentQuery.FieldByName("CEPINICIAL").Value
  QueryVerificaCEP.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  QueryVerificaCEP.Active = False
  QueryVerificaCEP.Active = True
  If QueryVerificaCEP.EOF Then
    bsShowMessage("CEP Inicial Inválido", "E")
    CanContinue = False
    Exit Sub
  Else
    CanContinue = True
  End If

  'se cepfinal existe
  QueryVerificaCEP.Clear
  QueryVerificaCEP.Add("SELECT CEP FROM LOGRADOUROS")
  QueryVerificaCEP.Add("WHERE  CEP = :CEPFINAL")
  QueryVerificaCEP.Add("  AND  HANDLE <> :HANDLE")
  QueryVerificaCEP.ParamByName("CEPFINAL").Value = CurrentQuery.FieldByName("CEPFINAL").Value
  QueryVerificaCEP.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  QueryVerificaCEP.Active = False
  QueryVerificaCEP.Active = True
  If QueryVerificaCEP.EOF Then
    bsShowMessage("CEP Final Inválido", "E")
    CanContinue = False
    Exit Sub
  Else
    CanContinue = True
  End If

  'se cepinicial naum pertence a outro distrito
  QueryVerificaCEP.Clear
  QueryVerificaCEP.Add("SELECT HANDLE FROM SAM_DISTRITOSANITARIOCEP DSCEP")
  QueryVerificaCEP.Add("WHERE  :CEPINICIAL >= DSCEP.CEPINICIAL")
  QueryVerificaCEP.Add("AND    :CEPINICIAL <= DSCEP.CEPFINAL")
  QueryVerificaCEP.Add("AND    DSCEP.HANDLE <> :HANDLE")
  QueryVerificaCEP.ParamByName("CEPINICIAL").Value = CurrentQuery.FieldByName("CEPINICIAL").Value
  QueryVerificaCEP.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  QueryVerificaCEP.Active = False
  QueryVerificaCEP.Active = True
  If QueryVerificaCEP.EOF Then
    CanContinue = True
  Else
    bsShowMessage("CEP Inicial pertencente a outra área", "E")
    CanContinue = False
    Exit Sub
  End If


  'se cepfinal naum pertence a outro distrito
  QueryVerificaCEP.Clear
  QueryVerificaCEP.Add("SELECT HANDLE FROM SAM_DISTRITOSANITARIOCEP DSCEP")
  QueryVerificaCEP.Add("WHERE  :CEPFINAL >= DSCEP.CEPINICIAL")
  QueryVerificaCEP.Add("AND    :CEPFINAL <= DSCEP.CEPFINAL")
  QueryVerificaCEP.Add("AND    DSCEP.HANDLE <> :HANDLE")
  QueryVerificaCEP.ParamByName("CEPFINAL").Value = CurrentQuery.FieldByName("CEPFINAL").Value
  QueryVerificaCEP.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  QueryVerificaCEP.Active = False
  QueryVerificaCEP.Active = True
  If QueryVerificaCEP.EOF Then
    CanContinue = True
  Else
    bsShowMessage("CEP Final pertencente a outra área", "E")
    CanContinue = False
    Exit Sub
  End If

  Set QueryVerificaCEP = Nothing

End Sub


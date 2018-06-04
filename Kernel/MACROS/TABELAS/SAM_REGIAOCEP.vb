'HASH: C680BCF550C7E06E598FBA7958111D26
'#Uses "*bsShowMessage"

'Macro SamRegiaoCep Shiba 07/10/2002

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim Cep1 As Long
  Dim Cep2 As Long
  Dim CEPIN As Long
  Dim CEPFI As Long


  CEPIN = Int(Mid(CurrentQuery.FieldByName("CEPINICIAL").AsString, 1, 5) + Mid(CurrentQuery.FieldByName("CEPINICIAL").AsString, 7, 3))
  CEPFI = Int(Mid(CurrentQuery.FieldByName("CEPfinal").AsString, 1, 5) + Mid(CurrentQuery.FieldByName("CEPfinal").AsString, 7, 3))

  If CEPFI < CEPIN Then
    bsShowMessage("O CEP final não pode ser menor que o CEP inicial!", "E")
    CanContinue = False
    Exit Sub
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT CEPINICIAL, CEPFINAL FROM SAM_REGIAOCEP WHERE HANDLE <> :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  While Not SQL.EOF

    Cep1 = Int(Mid(SQL.FieldByName("CEPINICIAL").AsString, 1, 5) + Mid(SQL.FieldByName("CEPINICIAL").AsString, 7, 3))
    Cep2 = Int(Mid(SQL.FieldByName("CEPFINAL").AsString, 1, 5) + Mid(SQL.FieldByName("CEPFINAL").AsString, 7, 3))

    If ((CEPIN >= Cep1) And (CEPIN <= Cep2)) Or _
         ((CEPFI >= Cep1) And (CEPFI <= Cep2)) Or _
         ((CEPIN <= Cep1) And (CEPFI >= Cep2)) Or _
         ((CEPIN >= Cep1) And (CEPFI <= Cep2)) Then

      bsShowMessage("O intervalo entre cep's não pode ser coincidente!", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
    SQL.Next
  Wend

  SQL.Active = False
  Set SQL = Nothing

End Sub


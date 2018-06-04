'HASH: E40FD812AD563B07094AAB8E2EB516CC
'Macro: SAM_ALCADAPAGTO

'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("VALORLIMITE").AsFloat >= 0 Then
    Dim qVerificaRegistro As Object
    Dim vNUlo As String
    Set qVerificaRegistro = NewQuery
    If InStr(SQLServer, "SQL") > 0 Then
      vNUlo = "ISNULL"
    End If
    If InStr(SQLServer, "DB2") > 0 Then
      vNUlo = "COALESCE"
    End If
    If InStr(SQLServer, "ORA") > 0 Then
      vNUlo = "NVL"
    End If
    If InStr(SQLServer, "CACHE") > 0 Then
      vNUlo = "NVL"
    End If

    qVerificaRegistro.Active = False
    qVerificaRegistro.Clear
    qVerificaRegistro.Add("SELECT HANDLE                                  ")
    qVerificaRegistro.Add("  FROM SAM_ALCADAPAGTO                         ")
    qVerificaRegistro.Add(" WHERE " + vNUlo + "(REGIMEATENDIMENTO,0) = :pREGIMEATENDIMENTO ")
    qVerificaRegistro.Add("   AND " + vNUlo + "(LOCALATENDIMENTO,0) = :pLOCALATENDIMENTO   ")
    qVerificaRegistro.Add("   AND HANDLE <> :pHANDLECORRENTE              ")
    qVerificaRegistro.Add("   AND VALORLIMITE = :VALORLIMITE              ")
    qVerificaRegistro.ParamByName("pREGIMEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
    qVerificaRegistro.ParamByName("pLOCALATENDIMENTO").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
    qVerificaRegistro.ParamByName("pHANDLECORRENTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qVerificaRegistro.ParamByName("VALORLIMITE").AsFloat = CurrentQuery.FieldByName("VALORLIMITE").AsFloat
    qVerificaRegistro.Active = True

    If (Not qVerificaRegistro.EOF) Then
        CanContinue = False
      bsShowMessage("Já há um registro com os mesmos dados.", "E")
    End If
    Set qVerificaRegistro = Nothing
  Else
    bsShowMessage("Não é permitido alçada com valor negativo!", "E")
    canconinue = False
  End If
End Sub


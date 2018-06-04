'HASH: 8C673C37FCD893C9C059AB1B201AB7EE
' Bruno (18/10/2004) SMS : 31109
'#Uses "*bsShowMessage"

Public Sub ALCADATIPO_OnExit()
  Dim qBuscaLimite As Object
  Set qBuscaLimite = NewQuery

  qBuscaLimite.Active = False
  qBuscaLimite.Clear
  qBuscaLimite.Add("SELECT VALORLIMITE FROM SAM_TIPOALCADA WHERE HANDLE = :pHANDLEALCADA")
  qBuscaLimite.ParamByName("pHANDLEALCADA").AsInteger = CurrentQuery.FieldByName("ALCADATIPO").AsInteger
  qBuscaLimite.Active = True

  If (CurrentQuery.State <> 1) Then
    CurrentQuery.FieldByName("LIMITE").AsFloat = qBuscaLimite.FieldByName("VALORLIMITE").AsFloat
  End If

  Set qBuscaLimite = Nothing
End Sub

Public Sub ALCADA_OnChange()
  If Not CurrentQuery.FieldByName("ALCADA").IsNull Then
    Dim Q As Object
    Set Q = NewQuery
    Q.Add("SELECT TEMLIMITE FROM Z_ALCADAS WHERE HANDLE = " + CurrentQuery.FieldByName("ALCADA").AsString)
    Q.Active = True
    If Q.FieldByName("TEMLIMITE").AsString = "S" Then
      LIMITE.Visible = True
    Else
      LIMITE.Visible = False
    End If
    Q.Active = False
    Set Q = Nothing
  End If
End Sub

Public Sub TABLE_AfterInsert()
  TABLE_AfterScroll
End Sub


Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("ALCADA").IsNull Then
    Dim Q As Object
    Set Q = NewQuery
    Q.Add("SELECT TEMLIMITE FROM Z_ALCADAS WHERE HANDLE = " + CurrentQuery.FieldByName("ALCADA").AsString)
    Q.Active = True
    If Q.FieldByName("TEMLIMITE").AsString = "S" Then
      LIMITE.Visible = True
    Else
      LIMITE.Visible = False
    End If
    Q.Active = False
    Set Q = Nothing

    ' Bruno (18/10/2004) SMS : 31109
    Dim qVerificaAlcada As Object
    Set qVerificaAlcada = NewQuery
    qVerificaAlcada.Active = False
    qVerificaAlcada.Clear
    qVerificaAlcada.Add("SELECT NOME FROM Z_ALCADAS WHERE HANDLE = :pHANDLEALCADA")
    qVerificaAlcada.ParamByName("pHANDLEALCADA").AsInteger = CurrentQuery.FieldByName("ALCADA").AsInteger
    qVerificaAlcada.Active = True

'    If (qVerificaAlcada.FieldByName("NOME").AsString = "ALCADA_PAGAMENTO") Then
'      ALCADATIPO.Visible = True
'    Else
'      ALCADATIPO.Visible = False
'    End If

  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT TEMLIMITE FROM Z_ALCADAS WHERE HANDLE = " + CurrentQuery.FieldByName("ALCADA").AsString)
  Q.Active = True
  If Q.FieldByName("TEMLIMITE").AsString = "S" And CurrentQuery.FieldByName("LIMITE").IsNull Then
    CanContinue = False
    bsShowMessage("Limite é obrigatório", "E")

  End If

  If Q.FieldByName("TEMLIMITE").AsString = "N" Then
    CurrentQuery.FieldByName("LIMITE").Clear
  End If
  
  Q.Active = False
  Set Q = Nothing

  ' Bruno (18/10/2004) SMS : 31109
  '   + Verificando se já existe algum registro com os mesmos dados
  Dim qVerificaRegistro As Object
  Set qVerificaRegistro = NewQuery
  qVerificaRegistro.Active = False
  qVerificaRegistro.Clear
  qVerificaRegistro.Add("SELECT                               ")
  qVerificaRegistro.Add("    HANDLE                           ")
  qVerificaRegistro.Add("  FROM                               ")
  qVerificaRegistro.Add("    Z_GRUPOUSUARIOALCADAS            ")
  qVerificaRegistro.Add("  WHERE                              ")
  qVerificaRegistro.Add("    ALCADA     = :pALCADA         AND")
  qVerificaRegistro.Add("    USUARIO    = :pUSUARIO        AND")
  qVerificaRegistro.Add("    HANDLE    <> :pHANDLECORRENTE    ")

'  If (CurrentQuery.FieldByName("ALCADATIPO").IsNull) Then
'    qVerificaRegistro.Add("    ALCADATIPO IS NULL")
'  Else
'    qVerificaRegistro.Add("    ALCADATIPO = :pALCADATIPO")
'  End If
  qVerificaRegistro.ParamByName("pALCADA").AsInteger = CurrentQuery.FieldByName("ALCADA").AsInteger
  qVerificaRegistro.ParamByName("pUSUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  qVerificaRegistro.ParamByName("pHANDLECORRENTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
'  If (Not CurrentQuery.FieldByName("ALCADATIPO").IsNull) Then
'    qVerificaRegistro.ParamByName("pALCADATIPO").AsInteger = CurrentQuery.FieldByName("ALCADATIPO").AsInteger
'  End If
  qVerificaRegistro.Active = True

  If (Not qVerificaRegistro.EOF) Then
    CanContinue = False
    bsShowMessage("Já existe um registro com os mesmos dados.", "E")
  End If

End Sub


'HASH: F07DA2341C4BF025D84B8F789338C56A
'#Uses "*bsShowMessage"
'#Uses "*bsShowMessageStyle"
Dim vsEncontrouPrincipal As Double
Dim vsChamandoMacroWebPart As String

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If Not CLng(SessionVar("SAM_MATRICULA")) > 0 Then
    bsShowMessage("Não existe um participante selecionado!", "E")
    CanContinue = False
  End If
End Sub



Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("ATENDIMENTO").Value = SessionVar("CLI_SUBJETIVO")
  CurrentQuery.FieldByName("DATAREGISTRO").AsDateTime = ServerNow
  CurrentQuery.FieldByName("DATAREGISTRO").AsDateTime = ServerDate
  CurrentQuery.FieldByName("MATRICULA").Value = SessionVar("SAM_MATRICULA")
End Sub





Public Sub TABLE_BeforePost(CanContinue As Boolean)
  vsEncontrouPrincipal = 0

  If CurrentQuery.FieldByName("MATRICULA").IsNull Then
    bsShowMessage("Participante não informado!", "E")
    CanContinue = False
    Exit Sub
  End If

  'If (CurrentQuery.FieldByName("EHCIDPRINCIPAL").AsString = "S") And (vsChamandoMacroWebPart = "S") Then
  If (CurrentQuery.FieldByName("EHCIDPRINCIPAL").AsString = "S")  Then
    Dim Q As Object
    Set Q = NewQuery

    Q.Add("Select COUNT(1) ACHOU                                         ")
    Q.Add("  FROM CLI_PACIENTEDIAGNOSTICO PD                             ")
    Q.Add("WHERE PD.EHCIDPRINCIPAL = 'S'                                 ")
    Q.Add("   And PD.ATENDIMENTO In (Select Sub.Handle                   ")
    Q.Add("                            FROM CLI_SUBJETIVO Sub            ")
    Q.Add("                           WHERE Sub.DATAENCERRAMENTO Is Null ")
    Q.Add("                             And Sub.Handle = :ATENDIMENTO)   ")
    Q.Add("   And PD.ATENDIMENTO = :ATENDIMENTO                          ")
    Q.Add("   AND PD.HANDLE <> :HANDLE                                   ")
    Q.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").Value
    Q.ParamByName("pHandle").Value = CurrentQuery.FieldByName("HANDLE").Value
    Q.Active = True

    vsEncontrouPrincipal = Q.FieldByName("ACHOU").AsInteger

     If (CurrentQuery.FieldByName("EHCIDPRINCIPAL").AsString = "S") And (vsEncontrouPrincipal > 0) Then
       AlterarCidPrincipal
     End If

  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "ALTERARCIDPRINCIPAL" Then
    AlterarCidPrincipal
  End If

End Sub


Public Sub AlterarCidPrincipal
    Dim Q As Object
    Set Q = NewQuery

    Q.Active = False
    Q.Add("UPDATE CLI_PACIENTEDIAGNOSTICO                             ")
    Q.Add("   SET EHCIDPRINCIPAL = 'N'                                ")
    Q.Add(" WHERE EHCIDPRINCIPAL = 'S'                                ")
    Q.Add("   And ATENDIMENTO In (SELECT Sub.Handle                   ")
    Q.Add("                         FROM CLI_SUBJETIVO Sub            ")
    Q.Add("                        WHERE Sub.DATAENCERRAMENTO Is Null ")
    Q.Add("                          And Sub.Handle = :ATENDIMENTO)   ")
    Q.Add("   And ATENDIMENTO = :ATENDIMENTO                          ")
    Q.Add("   AND HANDLE <> :HANDLE                                   ")
    Q.ParamByName("pAtendimento").Value = CurrentQuery.FieldByName("ATENDIMENTO").Value
    Q.ParamByName("pHandle").Value = CurrentQuery.FieldByName("HANDLE").Value
    Q.ExecSQL

    Set Q = Nothing
End Sub



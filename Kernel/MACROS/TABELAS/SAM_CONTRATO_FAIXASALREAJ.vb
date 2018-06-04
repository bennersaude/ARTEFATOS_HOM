'HASH: 00678A9D26681D4F07EE8C879189161A
'MACRO : SAM_CONTRATO_FAIXASALREAJ
' Alterações :
'  - 04/11/2004 : Bruno SMS 32871 -> Validação da vigência
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  If Not (CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull) Then
    COMPETENCIAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  ' Checando a validade da vigência

  If (CurrentQuery.FieldByName("COMPETENCIAINICIAL").Value > CurrentQuery.FieldByName("COMPETENCIAFINAL").Value) Then
    bsShowMessage("Data final não pode ser menor que a data inicial.", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim qVerificaVigencia As Object
  Set qVerificaVigencia = NewQuery

  qVerificaVigencia.Active = False
  qVerificaVigencia.Clear

  qVerificaVigencia.Add("SELECT                                                                                                     ")
  qVerificaVigencia.Add("    HANDLE                                                                                                 ")
  qVerificaVigencia.Add("  FROM                                                                                                     ")
  qVerificaVigencia.Add("    SAM_CONTRATO_FAIXASALREAJ                                                                              ")
  qVerificaVigencia.Add("  WHERE                                                                                                    ")
  qVerificaVigencia.Add("    HANDLE <> :pHANDLECURRENT AND                                                                          ")
  qVerificaVigencia.Add("    (                                                                                                      ")

  ' Quando a data final da vigência atual é nula
  If(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)Then
  	qVerificaVigencia.Add("    (COMPETENCIAFINAL IS NULL) OR                                                                        ")
  	qVerificaVigencia.Add("    (:pCOMPETENCIAINICIAL <= COMPETENCIAFINAL)                                                           ")
  Else
  	' Quando há alguma vigência com a data final nula já cadastrada
  	qVerificaVigencia.Add("    ((COMPETENCIAFINAL IS NULL) AND (:pCOMPETENCIAFINAL >= COMPETENCIAINICIAL)) OR                       ")

  	' Quando não há nenhuma vigência com a data final nula e a data final atual também não é nula
  	qVerificaVigencia.Add("      ((NOT(COMPETENCIAFINAL IS NULL)) AND                                                               ")
  	qVerificaVigencia.Add("        (                                                                                                ")

  	' Quando a data inicial corrente esta contida dentro de uma outra vigência já cadastrada
  	qVerificaVigencia.Add("         ((:pCOMPETENCIAINICIAL >= COMPETENCIAINICIAL) AND (:pCOMPETENCIAINICIAL <= COMPETENCIAFINAL)) OR")

  	' Quando a data final corrente esta contida dentro de uma outra vigência já cadastrada
  	qVerificaVigencia.Add("         ((:pCOMPETENCIAFINAL >= COMPETENCIAINICIAL) AND (:pCOMPETENCIAFINAL <= COMPETENCIAFINAL)) OR    ")

  	' Quando uma vigência inteira já cadastrada esta contida dentro da vigência atual
  	qVerificaVigencia.Add("         ((:pCOMPETENCIAFINAL >= COMPETENCIAFINAL) AND (:pCOMPETENCIAINICIAL <= COMPETENCIAINICIAL))     ")

  	qVerificaVigencia.Add("        )                                                                    ")
  	qVerificaVigencia.Add("      )                                                                      ")
  	qVerificaVigencia.ParamByName("pCOMPETENCIAFINAL").Value = CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime
  End If

  qVerificaVigencia.Add("    )                                                                          ")
  qVerificaVigencia.Add("    AND CONTRATO = :pCONTRATOCURRENT                                           ")
  qVerificaVigencia.ParamByName("pCOMPETENCIAINICIAL").Value = CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime
  qVerificaVigencia.ParamByName("pHANDLECURRENT").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaVigencia.ParamByName("pCONTRATOCURRENT").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger

  qVerificaVigencia.Active = True

  If(Not qVerificaVigencia.EOF)Then
  	bsShowMessage("Vigência inválida.", "E")
  	CanContinue = False
  End If

  Set qVerificaVigencia = Nothing
End Sub


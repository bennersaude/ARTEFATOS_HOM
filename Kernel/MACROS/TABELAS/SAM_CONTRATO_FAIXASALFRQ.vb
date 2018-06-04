'HASH: 733393C6665E9ABD708A8CFFE2A5E998
'MACRO : SAM_CONTRATO_FAIXASALFRQ
' Alterações :
'  - 04/11/2004 : Bruno SMS 32871
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  ' Julio - SMS 81728 - Início
  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    DATAFINAL.ReadOnly = True
  Else
    DATAFINAL.ReadOnly = False
  End If
  ' Julio - SMS 81728 - Fim
End Sub

Public Sub TABLE_AfterScroll()
  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    DATAFINAL.ReadOnly = True
  Else ' Julio - SMS 81728
    DATAFINAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  ' Checando a validade da vigência

  If (CurrentQuery.FieldByName("DATAINICIAL").Value > CurrentQuery.FieldByName("DATAFINAL").Value) Then
    bsShowMessage("Data final não pode ser menor que a data inicial.", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim qVerificaVigencia As Object
  Set qVerificaVigencia = NewQuery

  qVerificaVigencia.Active = False
  qVerificaVigencia.Clear

  qVerificaVigencia.Add("SELECT                                                                         ")
  qVerificaVigencia.Add("    HANDLE                                                                     ")
  qVerificaVigencia.Add("  FROM                                                                         ")
  qVerificaVigencia.Add("    SAM_CONTRATO_FAIXASALFRQ                                                   ")
  qVerificaVigencia.Add("  WHERE                                                                        ")
  qVerificaVigencia.Add("    HANDLE <> :pHANDLECURRENT AND                                              ")
  qVerificaVigencia.Add("    (                                                                          ")

  ' Quando a data final da vigência atual é nula
  If(CurrentQuery.FieldByName("DATAFINAL").IsNull)Then
  	qVerificaVigencia.Add("    (DATAFINAL IS NULL) OR                                                   ")
  	qVerificaVigencia.Add("    (:pDATAINICIAL <= DATAFINAL)                                             ")
  Else
  	' Quando há alguma vigência com a data final nula já cadastrada
  	qVerificaVigencia.Add("    ((DATAFINAL IS NULL) AND (:pDATAFINAL >= DATAINICIAL)) OR                ")

  	' Quando não há nenhuma vigência com a data final nula e a data final atual também não é nula
  	qVerificaVigencia.Add("      ((NOT(DATAFINAL IS NULL)) AND                                          ")
  	qVerificaVigencia.Add("        (                                                                    ")

  	' Quando a data inicial corrente esta contida dentro de uma outra vigência já cadastrada
  	qVerificaVigencia.Add("         ((:pDATAINICIAL >= DATAINICIAL) AND (:pDATAINICIAL <= DATAFINAL)) OR")

	  ' Quando a data final corrente esta contida dentro de uma outra vigência já cadastrada
  	qVerificaVigencia.Add("         ((:pDATAFINAL >= DATAINICIAL) AND (:pDATAFINAL <= DATAFINAL)) OR    ")

  	' Quando uma vigência inteira já cadastrada esta contida dentro da vigência atual
  	qVerificaVigencia.Add("         ((:pDATAFINAL >= DATAFINAL) AND (:pDATAINICIAL <= DATAINICIAL))     ")

  	qVerificaVigencia.Add("        )                                                                    ")
  	qVerificaVigencia.Add("      )                                                                      ")
  	qVerificaVigencia.ParamByName("pDATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  End If

  qVerificaVigencia.Add("    )                                                                          ")
  qVerificaVigencia.Add("    AND CONTRATO = :pCONTRATOCURRENT                                           ")
  qVerificaVigencia.ParamByName("pDATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  qVerificaVigencia.ParamByName("pHANDLECURRENT").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaVigencia.ParamByName("pCONTRATOCURRENT").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger

  qVerificaVigencia.Active = True

  If(Not qVerificaVigencia.EOF)Then
	bsShowMessage("Vigência inválida.", "E")
	CanContinue = False
  End If

  Set qVerificaVigencia = Nothing
End Sub


'HASH: 5F3B80CCF64EDA46C09E6F7F9C393156

'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  If VisibleMode Then
    CurrentQuery.FieldByName("VERSAOTISS").AsInteger = RecordHandleOfTable("TIS_VERSAO")
  End If
End Sub

Public Sub TABLE_AfterScroll()    	' SMS 104453 - Paulo Melo - 28/10/2008
  If WebMode Then
    GRAU.WebLocalWhere = "TABTIPOGRAU = 2"
  Else
    GRAU.LocalWhere = "TABTIPOGRAU = 2"
  End If

  If WebMode Then
    DENTE2.WebLocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
    REGIAO.WebLocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
  Else
    If Not VisibleMode Then
      DENTE2.LocalWhere = " VERSAOTISS = " + CurrentQuery.FieldByName("VERSAOTISS").AsString
      REGIAO.LocalWhere = " VERSAOTISS = " + CurrentQuery.FieldByName("VERSAOTISS").AsString
    Else
      DENTE2.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
      REGIAO.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
    End If
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("GRAU").IsNull Then Exit Sub

  Dim qAux As Object
  Set qAux = NewQuery
  qAux.Active = False
  qAux.Clear
  qAux.Add("SELECT DFC.HANDLE, CDT.DESCRICAO AS DENTE, RDT.DESCRICAO AS REGIAO ")
  qAux.Add("  FROM TIS_DENTEFACE      DFC                             		 ")
  qAux.Add("  LEFT JOIN TIS_CODIGODENTE    CDT On (CDT.Handle = DFC.DENTE2)	 ")
  qAux.Add("  LEFT JOIN TIS_REGIAODENTARIA RDT On (RDT.Handle = DFC.REGIAO)	 ")
  qAux.Add("  JOIN SAM_GRAU           GRA On (GRA.Handle = DFC.GRAU)  		 ")
  qAux.Add(" WHERE DFC.GRAU = " + CurrentQuery.FieldByName("GRAU").AsString 	  )
  qAux.Add("   AND DFC.HANDLE <> (:HANDLE)									 ") ' SMS 103195 - Danilo Raisi
  qAux.Add("   AND DFC.VERSAOTISS = :VERSAOTISS    							 ")
  qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger ' SMS 103195 - Danilo Raisi
  If Not VisibleMode Then
    qAux.ParamByName("VERSAOTISS").AsInteger = CurrentQuery.FieldByName("VERSAOTISS").AsString
  Else
    qAux.ParamByName("VERSAOTISS").AsInteger = RecordHandleOfTable("TIS_VERSAO")
  End If
  qAux.Active = True

  If Not qAux.EOF Then
    bsShowMessage("O Grau selecionado já foi utilizado para o dente """ + qAux.FieldByName("DENTE").AsString + """ e região dentária """ + qAux.FieldByName("REGIAO").AsString + """", "E")
    CanContinue = False
    Exit Sub
  End If

  qAux.Active = False
  qAux.Clear
  qAux.Add("SELECT DFC.HANDLE, CDT.DESCRICAO DENTE, RDT.DESCRICAO REGIAO, GRA.DESCRICAO GRAU ")
  qAux.Add("  FROM TIS_DENTEFACE       DFC								   				   ")
  qAux.Add("  LEFT JOIN TIS_CODIGODENTE     CDT ON (CDT.HANDLE = DFC.DENTE2)	   			   ")
  qAux.Add("  LEFT JOIN TIS_REGIAODENTARIA  RDT ON (RDT.HANDLE = DFC.REGIAO)	    		   ")
  qAux.Add("  JOIN SAM_GRAU			 GRA ON (GRA.HANDLE = DFC.GRAU)		   		  		   ")
  qAux.Add(" WHERE DFC.HANDLE <> (:HANDLE)  												   ")

  If CurrentQuery.FieldByName("DENTE2").AsInteger > 0 Then
    qAux.Add("   AND DFC.DENTE2 = (:DENTE)   										 	   ")
  End If

  If CurrentQuery.FieldByName("REGIAO").AsInteger > 0 Then
    qAux.Add("   AND DFC.REGIAO = (:REGIAO)										 		   ")
  End If

  qAux.Add("   AND DFC.FACEOCLUSAL = (:FACEOCLUSAL)										   ")
  qAux.Add("   AND DFC.FACELINGUAL = (:FACELINGUAL)										   ")
  qAux.Add("   AND DFC.FACEMESIAL = (:FACEMESIAL)										 	   ")
  qAux.Add("   AND DFC.FACEVESTIBULAR = (:FACEVESTIBULAR)									   ")
  qAux.Add("   AND DFC.FACEDISTAL = (:FACEDISTAL)										 	   ")
  qAux.Add("   AND DFC.FACEINCISAL = (:FACEINCISAL)										   ")
  qAux.Add("   AND DFC.PALATINA = (:PALATINA)										 		   ")
  qAux.Add("   AND DFC.VERSAOTISS = (:VERSAOTISS)                                            ")
  qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  If CurrentQuery.FieldByName("DENTE2").AsInteger > 0 Then
    qAux.ParamByName("DENTE").AsInteger = CurrentQuery.FieldByName("DENTE2").AsInteger
  End If

  If CurrentQuery.FieldByName("REGIAO").AsInteger > 0 Then
    qAux.ParamByName("REGIAO").AsInteger = CurrentQuery.FieldByName("REGIAO").AsInteger
  End If
  qAux.ParamByName("FACEOCLUSAL").AsString = CurrentQuery.FieldByName("FACEOCLUSAL").AsString
  qAux.ParamByName("FACELINGUAL").AsString = CurrentQuery.FieldByName("FACELINGUAL").AsString
  qAux.ParamByName("FACEMESIAL").AsString = CurrentQuery.FieldByName("FACEMESIAL").AsString
  qAux.ParamByName("FACEVESTIBULAR").AsString = CurrentQuery.FieldByName("FACEVESTIBULAR").AsString
  qAux.ParamByName("FACEDISTAL").AsString = CurrentQuery.FieldByName("FACEDISTAL").AsString
  qAux.ParamByName("FACEINCISAL").AsString = CurrentQuery.FieldByName("FACEINCISAL").AsString
  qAux.ParamByName("PALATINA").AsString = CurrentQuery.FieldByName("PALATINA").AsString
  qAux.ParamByName("VERSAOTISS").AsInteger = CurrentQuery.FieldByName("VERSAOTISS").AsInteger
  qAux.Active = True

  If Not qAux.EOF Then

    Dim vsMsg As String
    vsMsg = "Para configuração de Dente/Região/Face já existe grau cadastrado"

    If Not qAux.FieldByName("DENTE").IsNull Then
      vsMsg = vsMsg + Chr(13) + "Dente: " + qAux.FieldByName("DENTE").AsString
    End If

    If Not qAux.FieldByName("REGIAO").IsNull Then
      vsMsg = vsMsg + Chr(13) + "Região: " + qAux.FieldByName("REGIAO").AsString
    End If

    bsShowMessage(vsMsg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_NewRecord()

  If WebMode Then
    DENTE2.WebLocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
    REGIAO.WebLocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
  Else
    If Not VisibleMode Then
      DENTE2.LocalWhere = " VERSAOTISS = " + CurrentQuery.FieldByName("VERSAOTISS").AsString
      REGIAO.LocalWhere = " VERSAOTISS = " + CurrentQuery.FieldByName("VERSAOTISS").AsString
    Else
      DENTE2.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
      REGIAO.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO"))
    End If
  End If

End Sub

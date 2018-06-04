'HASH: EB39BA0B374AC3157E4FE88B48604EC3
'Tabela  TV_ANS002
'#Uses "*bsShowMessage"

Public Sub TABLE_UpdateRequired()

  If Not (((CurrentQuery.FieldByName("TABBUSCABENEFICIARIO").AsInteger = 1) And _
           (CurrentQuery.FieldByName("CCO").AsString = "" Or _
            CurrentQuery.FieldByName("CCODV").AsString = "")) Or _
           (CurrentQuery.FieldByName("TABBUSCABENEFICIARIO").AsInteger = 2 And _
            CurrentQuery.FieldByName("CODIGOANS").AsString = "")) Then

	  Dim qSelectBenef    As BPesquisa
	  Dim qSelecOperadora As BPesquisa
	  Dim chamaTV         As Boolean

	  Set qSelectBenef = NewQuery
	  Set qSelectOperadora = NewQuery

	  qSelectOperadora.Active = False
	  qSelectOperadora.Clear
	  qSelectOperadora.Add("   SELECT A.OPERADORA                                         ")
	  qSelectOperadora.Add("     FROM ANS_SIB_ROTINA A                                    ")
	  qSelectOperadora.Add("    WHERE A.HANDLE = (" + SessionVar("HANDLEROTINASIB") + ")  ")

	  qSelectOperadora.Active = True

	  chamaTV = False


	'Procura o beneficiário em uma rotina de conferência

	  qSelectBenef.Active = False
	  qSelectBenef.Clear
	  qSelectBenef.Add("  SELECT C.HANDLE                                          ")
	  qSelectBenef.Add("    FROM ANS_SIB_CADASTRO C                                ")
	  qSelectBenef.Add("   WHERE C.ROTINASIB IN (SELECT R.HANDLE                   ")
	  qSelectBenef.Add("                           FROM ANS_SIB_ROTINA R           ")
	  qSelectBenef.Add("                          WHERE R.TIPO = 3                 ")
	  qSelectBenef.Add("                            AND R.SITUACAO = 'P'           ")
	  qSelectBenef.Add("                            AND R.OPERADORA = :OPERADORA)  ")

	  If CurrentQuery.FieldByName("TABBUSCABENEFICIARIO").AsInteger = 1 Then
	    qSelectBenef.Add("     AND C.CCO = :CCO                                    ")
	    qSelectBenef.Add("     AND C.CCODV = :DV                                   ")
	    qSelectBenef.ParamByName("CCO").AsString = CurrentQuery.FieldByName("CCO").AsString
	    qSelectBenef.ParamByName("DV").AsString  = CurrentQuery.FieldByName("CCODV").AsString
	  Else
	    qSelectBenef.Add("     AND C.CODIGOBENEFANS = :CODIGOANS                   ")
	    qSelectBenef.ParamByName("CODIGOANS").AsString = CurrentQuery.FieldByName("CODIGOANS").AsString
	  End If
	    qSelectBenef.Add("     ORDER BY HANDLE DESC                                ")
	    qSelectBenef.ParamByName("OPERADORA").AsInteger = qSelectOperadora.FieldByName("OPERADORA").AsInteger
	    qSelectBenef.Active = True
	    qSelectBenef.First

	  If qSelectBenef.FieldByName("HANDLE").AsInteger <= 0 Then

	    qSelectBenef.Active = False
	    qSelectBenef.Clear
	    qSelectBenef.Add("  SELECT C.HANDLE                  ")
	    qSelectBenef.Add("    FROM ANS_SIB_CADASTRO C        ")
	    qSelectBenef.Add("   WHERE C.ROTINAENVIO IS NOT NULL ")

	    If CurrentQuery.FieldByName("TABBUSCABENEFICIARIO").AsInteger = 1 Then
	      qSelectBenef.Add("     AND C.CCO = :CCO            ")
	      qSelectBenef.Add("     AND C.CCODV = :DV           ")
	      qSelectBenef.ParamByName("CCO").AsString = CurrentQuery.FieldByName("CCO").AsString
	      qSelectBenef.ParamByName("DV").AsString  = CurrentQuery.FieldByName("CCODV").AsString
	    Else
	      qSelectBenef.Add("     AND C.CODIGOBENEFANS = :CODIGOANS  ")
	      qSelectBenef.ParamByName("CODIGOANS").AsString = CurrentQuery.FieldByName("CODIGOANS").AsString
	    End If
	    qSelectBenef.Add("       ORDER BY HANDLE DESC        ")
	    qSelectBenef.Active = True

	    If qSelectBenef.FieldByName("HANDLE").AsInteger <= 0 Then
	      BsShowMessage("Beneficiário não encontrado", "I")
	    Else
	      chamaTV = True
	    End If
	  Else
	    chamaTV = True
	  End If



	  If chamaTV And VisibleMode Then
	    Dim form As CSVirtualForm
	    Set form = NewVirtualForm

	    form.Caption = "Ajustar Beneficiáio"
	    form.TableName = "TV_ANS003"
	    form.Width = 550
	    form.Height = 750
	    form.TransitoryVars("HANDLEROTINACADASTRO") = qSelectBenef.FieldByName("HANDLE").AsInteger
	    form.TransitoryVars("HANDLEROTINASIB") = SessionVar("HANDLEROTINASIB")
	    form.Show

	    Set form = Nothing
	  End If

	  Set qSelectBenef = Nothing
	  Set qSelecOperadora = Nothing

  End If

End Sub

'HASH: 6C1333694A55F65864F7B17A22773E52
'#uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()

Dim SQL As Object
Dim qCampos As Object

Set qCampos = NewQuery
Set SQL = NewQuery

qCampos.Add("SELECT * FROM SFN_TESOURARIA_LANC")
qCampos.Add("WHERE HANDLE = :PLANC")
qCampos.ParamByName("PLANC").AsInteger = RecordHandleOfTable("SFN_TESOURARIA_LANC")
qCampos.Active = True

CurrentQuery.FieldByName("NUMERO").AsInteger = qCampos.FieldByName("NUMERO").AsInteger
CurrentQuery.FieldByName("DATA").AsDateTime = qCampos.FieldByName("DATA").AsDateTime
CurrentQuery.FieldByName("DATACONTABIL").AsDateTime = qCampos.FieldByName("DATACONTABIL").AsDateTime
CurrentQuery.FieldByName("DATACONFERENCIA").AsDateTime = qCampos.FieldByName("CONFERIDODATA").AsDateTime
CurrentQuery.FieldByName("TESOURARIA").AsString = qCampos.FieldByName("TESOURARIA").AsString
CurrentQuery.FieldByName("VALOR").AsString = qCampos.FieldByName("VALOR").AsString
CurrentQuery.FieldByName("HISTORICO").AsString = qCampos.FieldByName("HISTORICO").AsString
CurrentQuery.FieldByName("FAVORECIDO").AsString = qCampos.FieldByName("FAVORECIDO").AsString
CurrentQuery.FieldByName("CHEQUE").AsString = qCampos.FieldByName("CHEQUE").AsString

End Sub

Public Sub TABLE_AfterScroll()
Dim vsHtml As String


vsHtml = "@<html>										" + _
  		 "	  <head>									" + _
         "		<title>Tabela</title>					" + _
       	 "		<style Type='text/css'>					" + _
	     "      table.tabelafatdoc td					" + _
    	 "      {										" + _
         "          border: 1px solid black;			" + _
         "   	    margin: 0px;						" + _
	     "          padding: 2px;						" + _
    	 "       }										" + _
		 "       </style>								" + _
    	 "	</head>										" + _
    	 "	<body>										"



Dim SQL As Object
Dim qCampos As Object

Set qCampos = NewQuery
Set SQL = NewQuery

qCampos.Add("SELECT * FROM SFN_TESOURARIA_LANC")
qCampos.Add("WHERE HANDLE = :PLANC")
qCampos.ParamByName("PLANC").AsInteger = RecordHandleOfTable("SFN_TESOURARIA_LANC")
qCampos.Active = True

SQL.Add("SELECT DISTINCT F.HANDLE,             	")
SQL.Add("F.DATAVENCIMENTO,						")
SQL.Add("F.DATAEMISSAO,							")
SQL.Add("F.DATACONTABIL,						")
SQL.Add("F.BAIXADATA,							")
SQL.Add("F.NUMERO,								")
SQL.Add("F.VALOR,								")
SQL.Add("F.NATUREZA,							")
SQL.Add("F.BAIXAJURO,							")
SQL.Add("F.BAIXAMULTA,							")
SQL.Add("F.BAIXACORRECAO						")
SQL.Add("FROM SFN_FATURA_LANC FL,				")
SQL.Add("SFN_FATURA F							")
SQL.Add("WHERE FL.TESOURARIALANC = :PFAT And	")
SQL.Add("F.Handle = FL.FATURA					")
SQL.ParamByName("PFAT").AsInteger = qCampos.FieldByName("HANDLE").AsInteger
SQL.Active = True

If Not SQL.EOF Then

	vsHtml = vsHtml +    "	<b>Faturas</b>								" + _
		        		 "	<div style='overflow:auto;width:inherit'> <TABLE Class= 'tabelafatdoc'>				" + _
					     "      <tr>									" + _
				         "       	<td><b>Vencimento</b></td>			" + _
				         "       	<td><b>Emiss&atilde;o</b></td>		" + _
				         "       	<td><b>Data Contabil</b></td>		" + _
				         "       	<td><b>Baixa</b></td>				" + _
				         "       	<td><b>N&uacute;mero</b></td>		" + _
				         "       	<td><b>VALOR</b></td>				" + _
				         "       	<td><b>D/C</b></td>					" + _
				         "       	<td><b>Juro</b></td>				" + _
				         "       	<td><b>Multa</b></td>				" + _
				         "       	<td><b>Correção</b></td>			" + _
				         "   	</tr>									"


	While Not SQL.EOF

		vsHtml = vsHtml +   "   	<tr>																	" + _
							"       	<td>" + SQL.FieldByName("DATAVENCIMENTO").AsString 	+ "</td>		" + _
					   		"       	<td>" + SQL.FieldByName("DATAEMISSAO").AsString    	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("DATACONTABIL").AsString   	+ "</td>		" + _
			            	"       	<td>" + SQL.FieldByName("BAIXADATA").AsString      	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("NUMERO").AsString         	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("VALOR").AsString  			+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("NATUREZA").AsString	 	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("BAIXAJURO").AsString	 	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("BAiXAMULTA").AsString	 	+ "</td>		" + _
	         				"       	<td>" + SQL.FieldByName("BAIXACORRECAO").AsString	+ "</td>		" + _
	         				"		</tr>																	"


		SQL.Next

	Wend


	vsHtml = vsHtml + 	"</TABLE>"

End If

SQL.Active = False
SQL.Clear
SQL.Add("SELECT * FROM SFN_DOCUMENTO ")
SQL.Add("WHERE TESOURARIALANC = :PDOC")
SQL.ParamByName("PDOC").AsInteger = qCampos.FieldByName("HANDLE").AsInteger
SQL.Active = True


If Not SQL.EOF Then
	vsHtml = vsHtml +	"<br>											" + _
						"<b>Documentos</b>								" + _
					  	"<TABLE Class= 'tabelafatdoc'>					" + _
						"      <tr>										" + _
	         			"       	<td><b>Vencimento</b></td>			" + _
	         			"       	<td><b>Emiss&atilde;o</b></td>		" + _
	         			"       	<td><b>Baixa</b></td>				" + _
	         			"       	<td><b>N&uacute;mero</b></td>		" + _
	         			"			<td><b>Valor<b></td>   				" + _
	         			"       	<td><b>D/C</b></td>					" + _
	         			"       	<td><b>Juro</b></td>				" + _
	         			"       	<td><b>Multa</b></td>				" + _
	         			"       	<td><b>Desconto</b></td>			" + _
	         			"   	</tr>									"


	While Not SQL.EOF


	vsHtml = vsHtml +		"		<tr>																			" + _
							"       	<td>" + SQL.FieldByName("DATAVENCIMENTO").AsString  	+ "</td>			" + _
					   		"       	<td>" + SQL.FieldByName("DATAEMISSAO").AsString     	+ "</td>			" + _
			            	"       	<td>" + SQL.FieldByName("BAIXADATA").AsString       	+ "</td>			" + _
	         				"       	<td>" + SQL.FieldByName("NUMERO").AsString   			+ "</td>			" + _
	         				"       	<td>" + SQL.FieldByName("VALOR").AsString	 			+ "</td>			" + _
	         				"       	<td>" + SQL.FieldByName("NATUREZA").AsString	 		+ "</td>			" + _
	         				"       	<td>" + SQL.FieldByName("VALORJURO").AsString	 		+ "</td>			" + _
	         				"       	<td>" + SQL.FieldByName("VALORMULTA").AsString			+ "</td>		    " + _
	         				"       	<td>" + SQL.FieldByName("VALORDESCONTO").AsString		+ "</td>		    " + _
	         				"  		</tr>																		"


		SQL.Next
	Wend


vsHtml = vsHtml +		" 	</TABLE>	</div>"

End If


vsHtml = vsHtml +	"	</body>		" + _
 				  	"</html>		"



ROTFATDOC.Text = vsHtml
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


	On Error GoTo erro

	Dim Interface As Object

	Set Interface = CreateBennerObject("SFNTESOURARIA.Tesouraria")

	bsShowMessage(Interface.Estorno(CurrentSystem,RecordHandleOfTable("SFN_TESOURARIA_LANC"),CurrentQuery.FieldByName("HISTORICO").AsString), "I")

	Set Interface = Nothing

	Exit Sub


	erro:
		bsShowMessage(Error, "E")
		CanContinue = False

End Sub

'HASH: 2A5064C97370112E0A5BC6315E453FC5
'#Uses "*addXML"

Sub Main


    Dim viHandleNF As Long

    viHandleNF = CLng(SessionVar("HNOTA"))


    Dim qNota As BPesquisa
    Dim qSelec As BPesquisa
    Dim qDoc As BPesquisa

    Set qNota = NewQuery
    Set qSelec = NewQuery
    Set qDoc = NewQuery

    qNota.Clear
    qNota.Add("SELECT CONTAFINANCEIRA FROM SFN_NOTA  WHERE HANDLE = (:HANDLE)")
    qNota.ParamByName("HANDLE").AsInteger = viHandleNF
    qNota.Active = True

   qSelec.Clear
    qSelec.Add("SELECT N.HANDLE,  N.NUMERO, ")   ' N.SERIE,                                                      ")
    qSelec.Add("  N.DATAEMISSAO, N.VALOR, N.HANDLE                                                         ")
    qSelec.Add("   FROM SFN_NOTA N                                                                                        ")
    qSelec.Add("WHERE N.CONTAFINANCEIRA = :PCONTAFIN                                                ")
    qSelec.Add("      AND N.HANDLE NOT IN ( SELECT NOTA FROM SFN_NOTA_DOCUMENTO) ")
    qSelec.Add(" ORDER BY N.NUMERO")
    qSelec.ParamByName("PCONTAFIN").AsInteger = qNota.FieldByName("CONTAFINANCEIRA").AsInteger
    qSelec.Active = True

    qDoc.Clear
    qDoc.Add("SELECT D.HANDLE, D.NUMERO, ")
    qDoc.Add("               CASE WHEN D.NATUREZA = 'C' THEN 'Crédito' ")
    qDoc.Add("                         ELSE 'Débito' END NATUREZA, ")
    qDoc.Add("               D.DATAEMISSAO, ")
    qDoc.Add("               D.DATAVENCIMENTO, ")
    qDoc.Add("               D.VALOR, ")
    qDoc.Add("               TD.DESCRICAO TIPODOCUMENTO, ")
    qDoc.Add("               D.HANDLE ")
    qDoc.Add("   FROM SFN_DOCUMENTO D, ")
    qDoc.Add("              SFN_TIPODOCUMENTO TD ")
    qDoc.Add("WHERE D.CONTAFINANCEIRA = :PCONTAFIN ")
    qDoc.Add("      AND D.TIPODOCUMENTO = TD.HANDLE ")
    qDoc.Add("      AND D.CANCDATA IS NULL ")
    qDoc.Add("      AND D.VALOR > 0 ")
    qDoc.Add("      AND NOT EXISTS(SELECT ND.DOCUMENTO ")
    qDoc.Add("                                        FROM SFN_NOTA_DOCUMENTO ND ")
    qDoc.Add("                                     WHERE ND.DOCUMENTO = D.HANDLE) ")
    qDoc.ParamByName("PCONTAFIN").AsInteger = qNota.FieldByName("CONTAFINANCEIRA").AsInteger
    qDoc.Active = True


     Dim xmlNota As String
     Dim i As Integer
     Dim SQL As Object
     Set SQL = NewQuery
     SQL.Add("SELECT SERIE FROM SFN_SERIENOTA                                                                   ")
     SQL.Add("WHERE HANDLE IN (SELECT SERIE FROM SFN_NOTA WHERE HANDLE=:PNOTA)")
     SQL.ParamByName("PNOTA").AsInteger = viHandleNF 
     SQL.Active = True

     xmlNota = "<notas>"

	While Not qSelec.EOF
		xmlNota = xmlNota + "<nota>"
		For i = 0 To qSelec.FieldCount - 1
		       xmlNota = xmlNota + addXML( LCase( qSelec.Fields( i ).Name ), qSelec.Fields( i ).Name, qSelec )
		Next i
                                xmlNota = xmlNota +  "<serie>" + SQL.FieldByName("SERIE").Asstring + "</serie>"
		xmlNota = xmlNota + "</nota>"
		qSelec.Next
	Wend

    xmlNota = xmlNota + "</notas>"


    Dim xmlDoc As String

     xmlDoc = "<documentos>"

	While Not qDoc.EOF
		xmlDoc = xmlDoc + "<documento>"
		For i = 0 To qDoc.FieldCount - 1
		       xmlDoc = xmlDoc + addXML( LCase( qDoc.Fields( i ).Name ), qDoc.Fields( i ).Name, qDoc)
		Next i

		xmlDoc = xmlDoc + "</documento>"
		qDoc.Next
	Wend

    xmlDoc = xmlDoc + "</documentos>"

    ServiceVar("psXmlNotas") = xmlNota
    ServiceVar("psXmlDocumentos") = xmlDoc

End Sub

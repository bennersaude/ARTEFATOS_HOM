'HASH: 88768DFEC94EDFB47EDEBBC4487E508A
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim FiltroTexto As String
Dim FiltroPEGSituacaoInicial As Integer
Dim FiltroPEGSituacaoFinal As Integer

If(Not CurrentQuery.FieldByName("PEGSituacaoInicial").IsNull)Then
FiltroPEGSituacaoInicial =CurrentQuery.FieldByName("PEGSituacaoInicial").AsInteger
Else
FiltroPEGSituacaoInicial =1
End If

If(Not CurrentQuery.FieldByName("PEGSituacaoFinal").IsNull)Then
FiltroPEGSituacaoFinal =CurrentQuery.FieldByName("PEGSituacaoFinal").AsInteger
Else
FiltroPEGSituacaoFinal =4
End If



FiltroTexto ="Data Pagto Inicial:" +FiltroDataInicial +" - " +"Data Pgto Final:" +FiltroDataFinal


Select Case FiltroPEGSituacaoInicial
Case 1
FiltroTexto =FiltroTexto +"  Situacao PEG Inicial: 1.Em Digitação"

Case 2
FiltroTexto =FiltroTexto +"  Situacao PEG Inicial: 2.Em Conferência"

Case 3
FiltroTexto =FiltroTexto +"  Situacao PEG Inicial: 3.Pronto"

Case 4
FiltroTexto =FiltroTexto +"  Situacao PEG Inicial: 4.Faturado"
End Select

Select Case FiltroPEGSituacaoFinal
Case 1
FiltroTexto =FiltroTexto +"  Situacao PEG Final: 1.Em Digitação"

Case 2
FiltroTexto =FiltroTexto +"  Situacao PEG Final: 2.Em Conferência"

Case 3
FiltroTexto =FiltroTexto +"  Situacao PEG Final: 3.Pronto"

Case 4
FiltroTexto =FiltroTexto +"  Situacao PEG Final: 4.Faturado"
End Select

End Sub

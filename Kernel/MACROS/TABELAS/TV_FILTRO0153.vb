'HASH: 6E65EA2F41FB647C483F66C7C107F4AE
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim vbReexibirFiltro As Boolean
Dim vbAlgumFiltro As Boolean

Dim FiltroSolicitante As Integer
Dim FiltroLocalExecucao As Integer
Dim FiltroDataEmissao As Date
Dim FiltroAutorizacaoInicial As Integer
Dim FiltroAutorizacaoFinal As Integer



If(Not CurrentQuery.FieldByName("SOLICITANTE").IsNull)Then
FiltroSolicitante = CurrentQuery.FieldByName("SOLICITANTE").AsInteger
Else
FiltroSolicitante = -1
End If

If(Not CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull)Then
FiltroLocalExecucao = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
Else
FiltroLocalExecucao = -1
End If

If(Not CurrentQuery.FieldByName("AUTORIZACAOINICIAL").IsNull)Then
FiltroAutorizacaoInicial = CurrentQuery.FieldByName("AUTORIZACAOINICIAL").AsInteger
Else
FiltroAutorizacaoInicial = -1
End If

If(Not CurrentQuery.FieldByName("AUTORIZACAOFINAL").IsNull)Then
FiltroAutorizacaoFinal = CurrentQuery.FieldByName("AUTORIZACAOFINAL").AsInteger
Else
FiltroAutorizacaoFinal = -1
End If

If(Not CurrentQuery.FieldByName("DATAEMISSAO").IsNull)Then
FiltroDataEmissao = CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime
Else
FiltroDataEmissao = -1
  If (FiltroSolicitante + FiltroLocalExecucao > -2) And (FiltroAutorizacaoFinal + FiltroAutorizacaoInicial = -2) Then
   bsShowMessage("Por favor, ao preencher algum prestador preencha também a data de emissão.", "E")
   CanContinue = False
   Exit Sub
  End If
End If
If  (CurrentQuery.FieldByName("SOLICITANTE").IsNull And  CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull And CurrentQuery.FieldByName("AUTORIZACAOINICIAL").IsNull And CurrentQuery.FieldByName("AUTORIZACAOFINAL").IsNull And CurrentQuery.FieldByName("DATAEMISSAO").IsNull) Then
	    bsShowMessage("Por favor, preencha pelo menos um campo do filtro.", "E")
	    CanContinue = False
	    Exit Sub
	  End If



End Sub

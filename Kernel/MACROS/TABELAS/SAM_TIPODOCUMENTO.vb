'HASH: 718CB5905A41BAFC3003B08FDFBC02AF
'#Uses "*bsShowMessage

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("TIPODOCUMENTOBENEFICIARIO").AsString = "N") And _
	 (CurrentQuery.FieldByName("TIPODOCUMENTOPRESTADOR").AsString = "N") And _
	 (CurrentQuery.FieldByName("TIPODOCUMENTOADM").AsString = "N") And _
	 (CurrentQuery.FieldByName("TIPODOCUMENTOAUTORIZACAO").AsString = "N") Then
	bsShowMessage("Informe um tipo de documento.", "E")
	CanContinue = False
	Exit Sub
  End If

  If CurrentQuery.FieldByName("PADRAOVINCULOAUTORIZMONITORFAX").AsString = "S" Then
    Dim qPadraoVinculo As Object
    Set qPadraoVinculo = NewQuery
    qPadraoVinculo.Add("SELECT CODIGO,                             ")
    qPadraoVinculo.Add("       DESCRICAO                           ")
    qPadraoVinculo.Add("  FROM SAM_TIPODOCUMENTO                   ")
    qPadraoVinculo.Add(" WHERE PADRAOVINCULOAUTORIZMONITORFAX ='S' ")
    qPadraoVinculo.Add("   AND HANDLE <> :HANDLE                   ")
    qPadraoVinculo.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qPadraoVinculo.Active = True
    If Not qPadraoVinculo.EOF Then
        bsShowMessage("Já existe tipo de documento padrão, por favor desvincule o documento '" _
            + qPadraoVinculo.FieldByName("CODIGO").AsString + " - " _
            + qPadraoVinculo.FieldByName("DESCRICAO").AsString _
            + "' antes de vincular um novo padrão.", "E")
        CanContinue = False
    End If
  End If
End Sub

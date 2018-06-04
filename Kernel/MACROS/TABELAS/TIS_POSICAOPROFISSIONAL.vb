'HASH: D39D055A5BFE0916ABFBBB54C0C56BBB
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qGrau As Object
  Set qGrau = NewQuery

  qGrau.Clear
  qGrau.Add("SELECT HANDLE " + _
            "  FROM TIS_POSICAOPROFISSIONAL " + _
            " WHERE GRAU = :GRAU" + _
            "  AND VERSAOTISS = :VERSAOTISS" + _
            "  AND HANDLE <> :HANDLE")
  qGrau.Active = False
  qGrau.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
  qGrau.ParamByName("VERSAOTISS").AsInteger = RecordHandleOfTable("TIS_VERSAO")
  qGrau.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qGrau.Active = True

  If Not qGrau.EOF Then
	bsShowMessage("Não é possível cadastrar o mesmo grau para posições profissionais distintas na mesma versão da TISS", "E")
	CanContinue = False
	Exit Sub
  End If
End Sub

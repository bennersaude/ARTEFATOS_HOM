'HASH: DAD72C53D71BE55C1F4571B26A73F6A9


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		TIPODOCUMENTO.WebLocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
	ElseIf VisibleMode Then
		TIPODOCUMENTO.LocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		TIPODOCUMENTO.WebLocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
	ElseIf VisibleMode Then
		TIPODOCUMENTO.LocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
	End If
End Sub

Public Sub TABLE_UpdateRequired()
  If WebMode And Not(CurrentQuery.FieldByName("CONTRATOTPDEP").IsNull) Then
    Dim qContratoTpDep As Object
    Set qContratoTpDep = NewQuery

    qContratoTpDep.Add("SELECT TIPODEPENDENTE")
    qContratoTpDep.Add("FROM SAM_CONTRATO_TPDEP")
    qContratoTpDep.Add("WHERE HANDLE = :HCONTRATOTPDEP")
    qContratoTpDep.ParamByName("HCONTRATOTPDEP").Value = CurrentQuery.FieldByName("CONTRATOTPDEP").AsInteger
    qContratoTpDep.Active = True
    CurrentQuery.FieldByName("TIPODEPENDENTE").AsInteger = qContratoTpDep.FieldByName("TIPODEPENDENTE").AsInteger

    Set qContratoTpDep = Nothing
  End If
End Sub

'HASH: E09729270F2EDFD0B3D3AB23535E3CE7
'Macro: SAM_BENEFICIARIO_DOCENTREGUE
Option Explicit
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  Dim vLocalWhere As String
  vLocalWhere = GetSQLTipoDocumento

  If WebMode Then
	  TIPODOCUMENTO.WebLocalWhere = vLocalWhere
  ElseIf VisibleMode Then
  	  TIPODOCUMENTO.LocalWhere = vLocalWhere
  End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim vMeses As Integer

  MESESVALIDADE.Text = ""
  vMeses = DateDiff("m", CurrentQuery.FieldByName("DATAENTREGA").AsDateTime, CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime)

  If vMeses = 1 Then
    MESESVALIDADE.Text = "   1 Mês de Validade"
  ElseIf vMeses > 1 Then
    MESESVALIDADE.Text = "   " + CStr(vMeses) + " Meses de Validade"
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim vLocalWhere As String
  vLocalWhere = GetSQLTipoDocumento

  If WebMode Then
	  TIPODOCUMENTO.WebLocalWhere = vLocalWhere
  ElseIf VisibleMode Then
  	  TIPODOCUMENTO.LocalWhere = vLocalWhere
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then 'Inclusão
    Dim SQL As Object
    Dim Especifico As Object
	Dim TRF As Integer
    Set SQL = NewQuery
	Set Especifico = CreateBennerObject("Especifico.uEspecifico")
	TRF = 14
	Dim vJoin As String
	Dim vDBO As String
	Dim vCampos As String
	Dim vSQLX As String
	Dim vWhereX As String


	If (InStr(SQLServer, "MSSQL") > 0) Then
	  vDBO = "DBO."
	Else
	  vDBO = ""
	End If

    vSQLX = "SELECT X.MESESVALIDADE, X.EXIGEANEXO, " + vDBO + "GETDATABYINT(X.DIAVALIDOATE, X.MESVALIDOATE," + _
      "  CASE WHEN X.COMPETENCIA_FINAL < X.COMPETENCIA_INICIAL THEN 1 ELSE 0 END + :ANO, X.ANOSUBSEQUENTE) VALIDOATE" + _
      " FROM ("
	vCampos = "CD.MESESVALIDADE, D.EXIGEANEXO, VD.DIAINICIAL, VD.MESINICIAL, VD.DIAFINAL, VD.MESFINAL, VD.HANDLE, VD.DIAVALIDOATE, VD.MESVALIDOATE, VD.ANOSUBSEQUENTE," + _
      " CAST(CAST(VD.MESINICIAL As VARCHAR(2)) + " + vDBO + "BS_LPAD(CAST(VD.DIAINICIAL As VARCHAR(2)), 2, '0') AS INTEGER) COMPETENCIA_INICIAL," + _
      " CAST(CAST(VD.MESFINAL As VARCHAR(2)) + " + vDBO + "BS_LPAD(CAST(VD.DIAFINAL As VARCHAR(2)), 2, '0') AS INTEGER) COMPETENCIA_FINAL"
	vJoin = " JOIN SAM_TIPODOCUMENTO D ON D.HANDLE = CD.TIPODOCUMENTO " + _
	  " LEFT OUTER JOIN SAM_CONTRATO_TPDEP_DOC_VALID VD ON VD.TIPODOCUMENTODEPENDENTE = CD.HANDLE"
    vWhereX = ") X WHERE ((X.HANDLE IS NULL) OR (:DATAENTREGA BETWEEN " + vDBO + "GETDATABYINT(X.DIAINICIAL, X.MESINICIAL, :ANO, 'N')" + _
      " AND " + vDBO + "GETDATABYINT(X.DIAFINAL, X.MESFINAL, CASE WHEN X.COMPETENCIA_FINAL < X.COMPETENCIA_INICIAL THEN 1 ELSE 0 END + :ANO, 'N')))"

	SQL.Clear
   	SQL.Add(GetSQLDocumento(vCampos, " AND CD.TIPODOCUMENTO = :TIPODOCUMENTO ", vJoin, vSQLX, vWhereX))
    SQL.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    SQL.ParamByName("TIPODOCUMENTO").Value = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
    SQL.ParamByName("DATAENTREGA").Value = CurrentQuery.FieldByName("DATAENTREGA").AsDateTime
    SQL.ParamByName("ANO").Value = Year(CurrentQuery.FieldByName("DATAENTREGA").AsDateTime)
    SQL.Active = True

    If (CurrentQuery.FieldByName("DATAVALIDADE").IsNull) And (Not SQL.EOF) Then
        If (SQL.FieldByName("MESESVALIDADE").IsNull) And (SQL.FieldByName("VALIDOATE").IsNull) Then
          CurrentQuery.FieldByName("DATAVALIDADE").Clear
        Else
          If Year(SQL.FieldByName("VALIDOATE").AsDateTime) > 1900 Then
            CurrentQuery.FieldByName("DATAVALIDADE").Value = SQL.FieldByName("VALIDOATE").AsDateTime
          Else
            CurrentQuery.FieldByName("DATAVALIDADE").Value = _
                                   DateAdd("m", SQL.FieldByName("MESESVALIDADE").AsInteger, _
                                   CurrentQuery.FieldByName("DATAENTREGA").AsDateTime)
          End If
        End If
    End If

    If ((SQL.FieldByName("EXIGEANEXO").Value = "S") And (CurrentQuery.FieldByName("IMAGEMDOCUMENTO").IsNull) And (Especifico.Cliente(CurrentSystem) <> TRF)) Then
	  bsShowMessage("Este tipo de documento exige que seja informado um anexo", "E")
      CanContinue = False
      Exit Sub
    End If

    Set SQL = Nothing
  End If

  If (Not CurrentQuery.FieldByName("DATAVALIDADE").IsNull) And (CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime <= _
      CurrentQuery.FieldByName("DATAENTREGA").AsDateTime) Then
    bsShowMessage("A data de validade deve ser maior que a data de entrega", "E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.InInsertion Or CurrentQuery.InEdition) And (CurrentQuery.FieldByName("DATAENTREGA").AsDateTime > ServerNow) Then
  	bsShowMessage("A data de entrega não deve ser maior que a data atual!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Function GetSQLDocumento(campos As String, Optional ByVal filtroAdicional As String = "", _
  Optional ByVal joinAdicional = "", Optional ByVal SQLX = "", Optional ByVal WhereX = "") As String
   Dim vFiltro As String

   If Trim(filtroAdicional) <> "" Then
     vFiltro = "B.HANDLE = :BENEFICIARIO" + " " + filtroAdicional
   Else
     vFiltro = "B.HANDLE = " + CurrentQuery.FieldByName("BENEFICIARIO").AsString
   End If

   GetSQLDocumento = SQLX + "SELECT " + campos + " FROM SAM_BENEFICIARIO B "
   GetSQLDocumento = GetSQLDocumento + "  JOIN SAM_CONTRATO_TPDEP_DOC CD ON CD.CONTRATOTPDEP = B.TIPODEPENDENTE"
   GetSQLDocumento = GetSQLDocumento + " " + joinAdicional
   GetSQLDocumento = GetSQLDocumento + " WHERE " + vFiltro
   GetSQLDocumento = GetSQLDocumento + "  AND CD.CONTRATO = B.CONTRATO" + WhereX
End Function

Public Function GetSQLTipoDocumento As String
  GetSQLTipoDocumento = "HANDLE IN (" + GetSQLDocumento("CD.TIPODOCUMENTO") + ")"
End Function

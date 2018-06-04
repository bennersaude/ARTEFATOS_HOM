'HASH: 8D392EB78A19754C6F5B3E638CB059CC
'MACRO TABELA: SAM_MEMBRO_ESPECIALIDADE

Dim vCondicao As String

Public Sub ESPECIALIDADE_OnPopup(ShowPopup As Boolean)
	UpdateLastUpdate("SAM_ESPECIALIDADE")

	Dim SQL As Object
	Dim qCORPOCLINICO As Object
	Dim vData As Variant
	Set qCORPOCLINICO = NewQuery

	qCORPOCLINICO.Add("SELECT ENTIDADE,PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE HANDLE=:HANDLE")

	qCORPOCLINICO.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CORPOCLINICO").Value
	qCORPOCLINICO.Active = True

	Set SQL = NewQuery

	SQL.Add("SELECT * ")
	SQL.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE A")
	SQL.Add(" WHERE A.PRESTADOR = :ENTIDADE")
	SQL.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
	SQL.Add("   AND EXISTS (SELECT * ")
	SQL.Add("                 FROM SAM_PRESTADOR_ESPECIALIDADE  B")
	SQL.Add("                WHERE B.DATAINICIAL <= :DATA AND (B.DATAFINAL >= :DATA OR B.DATAFINAL IS NULL)")
	SQL.Add("                  AND B.ESPECIALIDADE = A.ESPECIALIDADE")
	SQL.Add("                  AND B.PRESTADOR = :PRESTADOR)")

	SQL.ParamByName("ENTIDADE").Value = qCORPOCLINICO.FieldByName("ENTIDADE").AsInteger
	SQL.ParamByName("PRESTADOR").Value = qCORPOCLINICO.FieldByName("PRESTADOR").AsInteger
	SQL.ParamByName("DATA").Value = ServerDate
	SQL.Active = True

	vData = ServerDate
	vCondicao = ""
	vCondicao = vCondicao + "SAM_ESPECIALIDADE.HANDLE"
	vCondicao = vCondicao + " IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE"
	vCondicao = vCondicao + "      WHERE ESPECIALIDADE     = " + SQL.FieldByName("ESPECIALIDADE").AsInteger

	While Not SQL.EOF
		vCondicao = vCondicao + "       OR ESPECIALIDADE      = " + SQL.FieldByName("ESPECIALIDADE").AsInteger

		SQL.Next
	Wend

	vCondicao = vCondicao + "  )"

	ESPECIALIDADE.LocalWhere = vCondicao

	Set SQL = Nothing
	Set qCORPOCLINICO = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If Not VisibleMode Then
		ESPECIALIDADE_OnPopup(False)
	End If
End Sub

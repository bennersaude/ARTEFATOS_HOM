'HASH: 93C6343B1C2C635C2D6E8A784548F756
 

Public Sub TABLE_AfterScroll()
	Dim SQL As BPesquisa
	Set SQL = NewQuery

	SQL.Add("SELECT B.NOME                          ")
	SQL.Add("   FROM SAM_BENEFICIARIO B             ")
	SQL.Add("        JOIN SAM_BENEFICIARIO_MOD BM   ")
	SQL.Add("        ON (B.HANDLE = BM.BENEFICIARIO)")
	SQL.Add("  WHERE BM.HANDLE = :HANDLE               ")

	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_MOD")

	SQL.Active = True

	BENEFICIARIO.Text = "Beneficiário: " + SQL.FieldByName("NOME").AsString

	CurrentQuery.Edit

	SQL.Active = False

	SQL.Clear

	SQL.Add("SELECT M.HANDLE									  ")
	SQL.Add("  FROM SAM_BENEFICIARIO_MOD BM						  ")
	SQL.Add("  JOIN SAM_CONTRATO_MOD CM ON (BM.MODULO = CM.HANDLE)")
	SQL.Add("  JOIN SAM_MODULO M ON (CM.MODULO = M.HANDLE)		  ")
	SQL.Add(" WHERE (BM.HANDLE = :HANDLE)						  ")


	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_MOD")

	SQL.Active = True

	CurrentQuery.FieldByName("MODULOORIGEM").AsInteger = SQL.FieldByName("HANDLE").AsInteger

	SQL.Active = False

	SQL.Clear

	SQL.Add("SELECT CM.PLANO              ")
	SQL.Add("  FROM SAM_CONTRATO_MOD CM")
	SQL.Add("       JOIN SAM_BENEFICIARIO_MOD BM")
	SQL.Add("       ON (CM.CONTRATO = BM.CONTRATO)")
	SQL.Add(" WHERE BM.HANDLE = :HANDLE")

    SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_MOD")

    SQL.Active = True

    CurrentQuery.FieldByName("PLANOORIGEM").AsInteger = SQL.FieldByName("PLANO").AsInteger

	SQL.Active = False

    SQL.Clear

    Dim qModReg As Object

    Set qModReg = NewQuery

    qModReg.Add("SELECT RMS.NOVAREGULAMENTACAO")
   	qModReg.Add("FROM SAM_REGISTROMS RMS")
    qModReg.Add("JOIN SAM_CONTRATO_MOD Mod On (RMS.Handle = Mod.REGISTROMS)")
	qModReg.Add("JOIN SAM_BENEFICIARIO_MOD BM ON (BM.MODULO = Mod.HANDLE)")
	qModReg.Add("WHERE BM.HANDLE = :HANDLE")

	qModReg.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_MOD")

	qModReg.Active = True

	SQL.Add("SELECT MIGRACAONREGREG,        ")
    SQL.Add("       MOTIVOMIGRACAO          ")
    SQL.Add("FROM SAM_PARAMETROSBENEFICIARIO")

    SQL.Active = True

    If qModReg.FieldByName("NOVAREGULAMENTACAO").AsString = "S" Then
        CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger = SQL.FieldByName("MOTIVOMIGRACAO").AsInteger
    Else
    	CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger = SQL.FieldByName("MIGRACAONREGREG").AsInteger
    End If


	Set SQL = Nothing
	Set qModReg = Nothing
End Sub

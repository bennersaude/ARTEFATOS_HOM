'HASH: BDE52E449B57893C18FA3F9375EF94DA
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T1115" Then
			MUNICIPIO.ReadOnly = True
			ESTADO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim DATA As Date
  Dim SQL As Object
  Set SQL = NewQuery

  ' SMS 73796 - Julio - 15/12/2006 - Ínicio
  SQL.Active = False
  SQL.Clear
  SQL.Add(" SELECT HANDLE                    ")
  SQL.Add("   FROM SAM_FERIADOMUNICIPAL      ")
  SQL.Add("  WHERE DATA      =  :QDATA       ")
  SQL.Add("    AND ESTADO    =  :QESTADO     ")
  SQL.Add("    AND MUNICIPIO =  :QMUNICIPIO  ")
  SQL.Add("    AND HANDLE    <> :QHANDLE     ")
  SQL.ParamByName("QDATA").AsDateTime  = CurrentQuery.FieldByName("DATA").AsDateTime
  SQL.ParamByName("QESTADO").AsInteger = CurrentQuery.FieldByName("ESTADO").AsInteger
  SQL.ParamByName("QMUNICIPIO").AsInteger = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
  SQL.ParamByName("QHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If (Not SQL.EOF) Then
    bsShowMessage("Já existe outro feriado cadastrado nesta data !", "E")
    CanContinue = False
    Exit Sub
  End If
  ' SMS 73796 - Julio - 15/12/2006 - Fim

  If CurrentQuery.FieldByName("BANCARIO").AsString = "N" Then
    DATA = DateValue(Str(DatePart("d", CurrentQuery.FieldByName("DATA").AsDateTime)) + "/" + _
           Str(DatePart("m", CurrentQuery.FieldByName("DATA").AsDateTime)) + "/" + _
           Str(DatePart("yyyy", CurrentQuery.FieldByName("DATA").AsDateTime)))
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT * FROM CLI_AGENDA ")
    SQL.Add(" WHERE DATAMARCADA = :DATA")
    SQL.Add("   AND DATADESMARCACAO IS NULL")
    SQL.ParamByName("DATA").AsDateTime = DATA
    SQL.Active = True
    If Not SQL.EOF Then
      bsShowMessage("Já existem consultas marcadas para esse dia!", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set SQL = Nothing
End Sub


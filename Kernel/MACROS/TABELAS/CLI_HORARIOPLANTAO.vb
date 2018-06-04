'HASH: B53CD9FF90E1C5F885B7493075A946A7
'CLI_HORARIOPLANTAO
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vEscala As Long
  Dim Ok As Boolean
  Dim ClinicaDLL As Object
  'início sms 38445 - Edilson.Castro - 10/06/2005
  'verificando se não há algum plantão com vigência cruzada
  Dim QVigenciaCruzada As Object

  Set QVigenciaCruzada = NewQuery

  'Condição incluída na SMS 62412 - 24.05.2006 - Incluido sinal de = na clausula e na mensagem SMS 63415(Item 3) - 12.06.2006
  If (CurrentQuery.FieldByName("HORAFINAL").AsDateTime <= CurrentQuery.FieldByName("HORAINICIAL").AsDateTime) Then
    bsShowMessage("A Hora Final não pode ser menor ou igual que a Hora Inicial.", "E")
    CanContinue = False
    Exit Sub
  End If
  'SMS 62412

  QVigenciaCruzada.Active = False
  QVigenciaCruzada.Clear
  QVigenciaCruzada.Add("SELECT 1                              ")
  QVigenciaCruzada.Add("  FROM CLI_HORARIOPLANTAO             ")
  QVigenciaCruzada.Add(" WHERE HANDLE <> :HANDLE              ")
  QVigenciaCruzada.Add("   AND CLINICA = :CLINICA             ")
  QVigenciaCruzada.Add("   AND RECURSO = :RECURSO             ")
  QVigenciaCruzada.Add("   AND ESPECIALIDADE = :ESPECIALIDADE ")
  QVigenciaCruzada.Add("   AND DATA = :DATA                   ")
  QVigenciaCruzada.Add("   AND (((HORAINICIAL <= :HORAINICIAL) AND (HORAFINAL   >= :HORAINICIAL)) ")
  QVigenciaCruzada.Add("     OR ((HORAINICIAL >  :HORAINICIAL) AND (HORAINICIAL <= :HORAFINAL)))  ")
  QVigenciaCruzada.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QVigenciaCruzada.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  QVigenciaCruzada.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  QVigenciaCruzada.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  QVigenciaCruzada.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATA").AsDateTime
  QVigenciaCruzada.ParamByName("HORAINICIAL").AsDateTime = CurrentQuery.FieldByName("HORAINICIAL").AsDateTime
  QVigenciaCruzada.ParamByName("HORAFINAL").AsDateTime = CurrentQuery.FieldByName("HORAFINAL").AsDateTime
  QVigenciaCruzada.Active = True

  If Not QVigenciaCruzada.EOF Then
    bsShowMessage("Existe um horário cadastrado para este recurso com vigência cruzada com a do horário desejado!", "E")
    CanContinue = False
    Exit Sub
  End If

  Set QVigenciaCruzada = Nothing
  'fim sms 38445

  Set ClinicaDLL = CreateBennerObject("CliClinica.Agenda")
  Ok = ClinicaDLL.TemEscalaHorario(CurrentSystem, CurrentQuery.FieldByName("DATA").AsDateTime, _
       CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, _
       CurrentQuery.FieldByName("RECURSO").AsInteger, _
       CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
       vEscala)
  If Not Ok Then
    CanContinue = False
    bsShowMessage("Período informado não pertence a nenhuma escala cadastrada para o prestador.", "E")
    Exit Sub
  End If
  Ok = ClinicaDLL.TemEscalaHorario(CurrentSystem, CurrentQuery.FieldByName("DATA").AsDateTime, _
       CurrentQuery.FieldByName("HORAFINAL").AsDateTime, _
       CurrentQuery.FieldByName("RECURSO").AsInteger, _
       CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
       vEscala)
  If Not Ok Then
    CanContinue = False
    bsShowMessage("Hora final está fora da escala!", "E")
    Exit Sub
  End If

  'SMS 59969 - Marcelo Barbosa - 29/03/2006
'	If CurrentQuery.State = 3 Then
'	  If InStr(SQLServer, "MSSQL") Then
'	    If Not CurrentQuery.FieldByName("HORAINICIAL").IsNull Then
'	      CurrentQuery.FieldByName("HORAINICIAL").AsDateTime =  Format("01/01/1900 " + Format(CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, "HH:MM:SS"), "MM/DD/YYYY HH:MM:SS")
'	    End If
'	    If Not CurrentQuery.FieldByName("HORAFINAL").IsNull Then
'	      CurrentQuery.FieldByName("HORAFINAL").AsDateTime =  Format("01/01/1900 " + Format(CurrentQuery.FieldByName("HORAFINAL").AsDateTime, "HH:MM:SS"), "MM/DD/YYYY HH:MM:SS")
'	    End If
'	  End If

'	  If InStr(SQLServer, "ORACLE") Then
'	    If Not CurrentQuery.FieldByName("HORAINICIAL").IsNull Then
'	      CurrentQuery.FieldByName("HORAINICIAL").AsDateTime =  Format("01/01/1900 " + Format(CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, "HH:MM:SS"), "DD/MM/YYYY HH:MM:SS")
'	    End If
'	    If Not CurrentQuery.FieldByName("HORAFINAL").IsNull Then
'	      CurrentQuery.FieldByName("HORAFINAL").AsDateTime =  Format("01/01/1900 " + Format(CurrentQuery.FieldByName("HORAFINAL").AsDateTime, "HH:MM:SS"), "DD/MM/YYYY HH:MM:SS")
'	    End If
'	  End If

'	 If InStr(SQLServer, "DB2") Then
'	    If Not CurrentQuery.FieldByName("HORAINICIAL").IsNull Then
'	       CurrentQuery.FieldByName("HORAINICIAL").AsDateTime =  Format("1900-01-01 " + Format(CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, "HH:MM:SS"), "YYYY-MM-DD HH:MM:SS")
'	    End If
'	    If Not CurrentQuery.FieldByName("HORAFINAL").IsNull Then
'	      CurrentQuery.FieldByName("HORAFINAL").AsDateTime =  Format("1900-01-01 " + Format(CurrentQuery.FieldByName("HORAFINAL").AsDateTime, "HH:MM:SS"), "YYYY-MM-DD HH:MM:SS")
'	    End If
'	 End If
'	End If
	'FIM - SMS 59969

  Set ClinicaDLL = Nothing
End Sub


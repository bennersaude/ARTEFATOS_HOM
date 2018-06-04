'HASH: 98D264AFE443926D864F1EDA98C6D107

Public Sub CALCULAR_OnClick()
  Set obj = CreateBennerObject("BSMed001.EspirometriaCoeficientes")
  obj.Exec(CurrentSystem)
  Set obj = Nothing

End Sub

Public Sub TABLE_AfterInsert()
  Dim Sql As Object, PRONTUARIO As Object
  Dim ALTURA2
  Dim IDADE2 As Integer
  Dim Anos As Long, Meses As Long, Dias As Long

  Set Sql = NewQuery
  Set PRONTUARIO = CreateBennerObject("CLIPRONTUARIO.Rotinas")
  Sql.Active = False
  Sql.Clear
  Sql.Add "SELECT  A.IDADE IDADE, A.ALTURA ALTURA, B.DATANASCIMENTO DATANASCIMENTO FROM MS_PACIENTES A, SAM_MATRICULA B WHERE A.MATRICULA = B.HANDLE AND A.HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES"))
  Sql.Active = True
  ALTURA2 = Sql.FieldByName("ALTURA").Value
  PRONTUARIO.Idade(CurrentSystem, Sql.FieldByName("DATANASCIMENTO").AsDateTime, Dias, Meses, Anos)
  IDADE2 = Anos
  If Not Sql.EOF Then
    CurrentQuery.FieldByName("IDADE").Value = IDADE2
    CurrentQuery.FieldByName("ALTURA").Value = ALTURA2
  End If
  Set Sql = Nothing
  Set Sql = NewQuery

End Sub

Public Sub TABLE_AfterPost()
  Dim Sql As Object, PRONTUARIO As Object
  Dim ALTURA2
  Dim IDADE2 As Integer, MATRICULA As Integer, TipoPaciente As Integer, NumHandle As Integer, HANDLE As Integer
  Dim Anos As Long, Meses As Long, Dias As Long

  If VisibleMode Then
    CALCULAR.Visible = True
    Set Sql = NewQuery
    Set PRONTUARIO = CreateBennerObject("CLIPRONTUARIO.Rotinas")

    Sql.Active = False
    Sql.Clear
    Sql.Add("SELECT  A.HANDLE HANDLE, A.IDADE IDADE, A.ALTURA ALTURA, A.MATRICULA MATRICULA, A.TIPOPACIENTE TIPOPACIENTE, B.DATANASCIMENTO DATANASCIMENTO FROM MS_PACIENTES A, SAM_MATRICULA B WHERE  A.HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES")))
    Sql.Active = True
    If Not(Sql.FieldByName("ALTURA").IsNull)Then
      ALTURA2 = Sql.FieldByName("ALTURA").Value
    Else
      ALTURA2 = 0
    End If
    If Not(Sql.FieldByName("IDADE").IsNull)Then
      IDADE2 = Sql.FieldByName("IDADE").AsInteger
    End If
    MATRICULA = Sql.FieldByName("MATRICULA").AsInteger
    TipoPaciente = Sql.FieldByName("TIPOPACIENTE").AsInteger
    HANDLE = Sql.FieldByName("HANDLE").AsInteger

    If Not Sql.EOF Then
      If(Not CurrentQuery.FieldByName("IDADE").IsNull)Then
      If(CurrentQuery.FieldByName("IDADE").AsInteger <>IDADE2)Then
      Sql.Active = False
      Sql.Clear
      Sql.Add "UPDATE MS_PACIENTES SET IDADE  = " + Str(CurrentQuery.FieldByName("IDADE").Value) + " WHERE HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES"))
      Sql.ExecSQL
    End If
  End If
  If(Not CurrentQuery.FieldByName("ALTURA").IsNull)Then
  If(CurrentQuery.FieldByName("ALTURA").Value <>ALTURA2)Then
  Sql.Active = False
  Sql.Clear
  Sql.Add "UPDATE MS_PACIENTES SET ALTURA = " + Str(CurrentQuery.FieldByName("ALTURA").Value) + " WHERE HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES"))
  Sql.ExecSQL
End If
End If

'  		If  TipoFuncionario =(1)Then
'  			Set Sql =Nothing
'    	    Set Sql =NewQuery

'  			Sql.Active =False
'     		Sql.Clear
'  			Sql.Add "SELECT  PESO, ALTURA, DATANASCIMENTO FROM DO_FUNCIONARIOS WHERE  HANDLE = " +CStr(Funcionario)
'  			Sql.Active =True
'  			If Not Sql.EOF Then
'  				If(Not Sql.FieldByName("ALTURA").IsNull)Then
'  					ALTURA =Calculo.ConverteValorString(Sql.FieldByName("ALTURA").Value)
'                End If
'                If (Not CurrentQuery.FieldByName("ALTURA").IsNull)Then

'    				If(CurrentQuery.FieldByName("ALTURA").Value <> ALTURA) Then
'    					SQL.Active =False
'     					Sql.Clear
'    					Sql.Add  "UPDATE DO_FUNCIONARIOS SET ALTURA = " +CStr(Calculo.ConverteValorString(CurrentQuery.FieldByName("ALTURA").Value))+" WHERE HANDLE = " +CStr(Funcionario)
'    					Sql.ExecSQL
'    				End If
'   			End If
'    		End If
'  		End If
End If


Set SQL = Nothing
'  	Set Calculo =Nothing
End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object, Calculo As Object
  Dim AlturaPaciente As Currency
  Set SQL = NewQuery
  'Set Calculo =CreateBennerObject("FP.CalculoFolha")
  If VisibleMode Then

    '      Sql.Active =False
    '      Sql.Clear
    '      Sql.Add("SELECT A.HANDLE HANDLE, A.IDADE IDADE, A.ALTURA ALTURA, A.MATRICULA MATRICULA, A.TIPOPACIENTE TIPOPACIENTE, B.DATANASCIMENTO DATANASCIMENTO FROM MS_PACIENTES A, SAM_MATRICULA B WHERE  A.HANDLE = " +Str(RecordHandleOfTable("MS_PACIENTES")))
    '  	  Sql.Active =True
    '  	  AlturaPaciente =0
    '  	  If Not(Sql.FieldByName("ALTURA").IsNull)Then
    '  	    AlturaPaciente =Str(Sql.FieldByName("ALTURA").Value)
    '      End If
    '      IdadePaciente =0
    '      If Not(Sql.FieldByName("IDADE").IsNull)Then
    '  	    IdadePaciente =Str(Sql.FieldByName("IDADE").Value)
    '      End If
    '    	If(CurrentQuery.FieldByName("IDADE").IsNull)And IdadePaciente =0 Then
    '	    	MsgBox "Campo Idade deve ser preenchido"
    '			CanContinue =False
    '		Else
    '		  If(CurrentQuery.FieldByName("IDADE").IsNull)Then
    '
    '  	        DataNasc =FormatDateTime2("YYYY",Sql.FieldByName("DATANASCIMENTO").AsString)
    '	        Data2 =FormatDateTime2("YYYY",Date)
    '	        Data4 =Data2 -DataNasc
    '	        CurrentQuery.FieldByName("IDADE").Value =Data4
    '	     End If

    '  		End If
    '    	If(CurrentQuery.FieldByName("ALTURA").IsNull)Then 'And AlturaPaciente =0
    ' 			MsgBox "Campo Altura deve ser preenchido"
    '			CanContinue =False
    '  		Else
    '  		  If(CurrentQuery.FieldByName("ALTURA").IsNull)Then
    '  		     CurrentQuery.FieldByName("ALTURA").Value = AlturaPaciente/10
    '  		  End If
    '  		End If
  End If
  Set SQL = Nothing
  '    Set Calculo =Nothing

End Sub

Public Sub TABLE_NewRecord()
  If VisibleMode Then
    CurrentQuery.FieldByName("FATOGERADOR").Value = 6
  End If
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  CALCULAR.Visible = False
  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If
End Sub


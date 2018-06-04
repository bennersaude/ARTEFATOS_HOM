'HASH: F3B0C6BDF061815C789A487C2E826ECB
Dim ALTURA2
Dim PESO2

Public Sub CIDS_OnClick()
  Set obj = CreateBennerObject("BSMed001.CidsAnamnese")
  obj.Exec(CurrentSystem)
  Set obj = Nothing
End Sub


Public Sub EXCLUIR_OnClick()
  Dim obj As Object
  Set obj = CreateBennerObject("BSMed001.ExcluirAnamnese")
  obj.Exec(CurrentSystem)
  Set obj = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object, PRONTUARIO As Object
  Dim IDADE2 As Integer
  Dim Anos As Long, Meses As Long, Dias As Long

  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add "SELECT  A.IDADE IDADE, A.PESO PESO, A.ALTURA ALTURA, B.DATANASCIMENTO DATANASCIMENTO FROM MS_PACIENTES A, SAM_MATRICULA B WHERE A.MATRICULA = B.HANDLE AND A.HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES"))
  SQL.Active = True
  ALTURA2 = SQL.FieldByName("ALTURA").Value
  PESO2 = SQL.FieldByName("PESO").Value
  If Not SQL.EOF Then
    CurrentQuery.FieldByName("ALTURA").Value = ALTURA2
    CurrentQuery.FieldByName("PESO").Value = PESO2
  End If
  Set SQL = Nothing
  Set SQL = NewQuery
End Sub

Public Sub TABLE_AfterPost()
  '  CIDS.Visible =True
  EXCLUIR.Visible = True
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
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
'  CIDS.Visible =False
EXCLUIR.Visible = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  '   Dim Calculo As Object
  Dim IMC
  Dim AlturaUtilizada, PesoUtilizado

  If VisibleMode Then
    '  		Set Calculo =CreateBennerObject("FP.CalculoFolha")
    Set SQL = NewQuery

    PesoUtilizado = o
    AlturaUtilizada = 0
    If(Not CurrentQuery.FieldByName("PESO").IsNull)Then
    PesoUtilizado = CurrentQuery.FieldByName("PESO").Value
  End If
  If(Not CurrentQuery.FieldByName("ALTURA").IsNull)Then
  AlturaUtilizada = CurrentQuery.FieldByName("ALTURA").Value
End If

SQL.Add "SELECT  * FROM MS_PACIENTES WHERE  HANDLE = " + Str(RecordHandleOfTable("MS_PACIENTES"))
SQL.Active = True
If Not SQL.EOF Then
  If PesoUtilizado = 0 Then
    If(Not SQL.FieldByName("PESO").IsNull)Then
    PesoUtilizado = SQL.FieldByName("PESO").Value
  End If
End If
If AlturaUtilizada = 0 Then
  If(Not SQL.FieldByName("ALTURA").IsNull)Then
  AlturaUtilizada = SQL.FieldByName("ALTURA").Value
End If

End If
If(PesoUtilizado >0)And(AlturaUtilizada >0)Then
IMC = PesoUtilizado / (AlturaUtilizada * AlturaUtilizada)
CurrentQuery.FieldByName("IMC").Value = IMC
End If

End If
CurrentQuery.FieldByName("PESO").Value = PesoUtilizado
CurrentQuery.FieldByName("ALTURA").Value = AlturaUtilizada

If(PesoUtilizado >0)And(PesoUtilizado <>PESO2)Then
SQL.Clear
SQL.Add("UPDATE MS_PACIENTES SET   PESO = :PESO WHERE HANDLE = :HANDLE")
SQL.ParamByName("HANDLE").DataType = ftInteger
SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("MS_PACIENTES")
SQL.ParamByName("PESO").DataType = ftFloat
SQL.ParamByName("PESO").Value = Str(CurrentQuery.FieldByName("PESO").Value)
'     	 	Sql.ParamByName("PESO").Value =CStr(Calculo.ConverteValorString(CurrentQuery.FieldByName("PESO").Value))
SQL.ExecSQL
End If

If(AlturaUtilizada >0)And(AlturaUtilizada <>ALTURA2)Then
SQL.Clear
SQL.Add("UPDATE MS_PACIENTES SET   ALTURA = :ALTURA WHERE HANDLE = :HANDLE")
SQL.ParamByName("HANDLE").DataType = ftInteger
SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("MS_PACIENTES")
SQL.ParamByName("ALTURA").DataType = ftFloat
ALTURA2 = CurrentQuery.FieldByName("ALTURA").Value
SQL.ParamByName("ALTURA").Value = ALTURA2
'     	 	Sql.ParamByName("ALTURA").Value =CStr(Calculo.ConverteValorString(CurrentQuery.FieldByName("ALTURA").Value))/100
SQL.ExecSQL
End If

If CurrentQuery.FieldByName("IMC").Value >0 Then
  CurrentQuery.FieldByName("LIMITEINFERIOR").Value = 20 * (AlturaUtilizada * AlturaUtilizada)
  CurrentQuery.FieldByName("LIMITESUPERIOR").Value = 25 * (AlturaUtilizada * AlturaUtilizada)
End If
Set Sql = Nothing
End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If
End Sub


Public Sub TABLE_AfterScroll()
  Dim Sql As Object
  Set Sql = NewQuery

  EXCLUIR.Enabled = True

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  EXCLUIR.Enabled = False
  Exit Sub
End If

End Sub


'HASH: 50C94F8A279F1515FF1F48BE3E8A336E
'MACRO: MS_PACIENTES

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vColunas As String, vCriterio As String, vCampo As String
  Dim vHandle As Long   ' SMS 95929 - Paulo Melo - 23/04/2008 - Era integer, dai dava overflow
  Dim Interface As Object
  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "BENEFICIARIO|NOME|MATRICULA"
  vCampos = "Beneficiário|Nome|Matrícula"

  vHandle = Interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 2, vCampos, vCriterio, "Beneficiário", False, BENEFICIARIO.Text)

  If vHandle > 0 Then  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = vHandle
End Sub

Public Sub BOTAOALTERARSETORCARGO_OnClick()
  Dim  BSMED001 As Object
  Set BSMED001 = CreateBennerObject("BSMED001.AlterarSetorCargo")
  BSMED001.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  RefreshNodesWithTable("MS_PACIENTES")
End Sub

Public Sub FICHAMEDICA_OnClick()
  Set obj = CreateBennerObject("BSMed001.Opcoes")
  obj.Exec(CurrentSystem)
  Set obj = Nothing
End Sub

Public Sub MATRICULA_OnPopup(ShowPopup As Boolean)
  Dim vColunas As String, vCriterio As String, vCampo As String
  Dim vHandle As Long   ' SMS 95929 - Paulo Melo - 23/04/2008 - Era integer, dai dava overflow
  Dim Interface As Object

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME|MATRICULA"
  vCampos = "Nome|Matrícula"
  vCriterio = "NOT EXISTS (SELECT 1 FROM SAM_BENEFICIARIO WHERE MATRICULA=SAM_MATRICULA.HANDLE)"

  vHandle = Interface.Exec(CurrentSystem, "SAM_MATRICULA", vColunas, 1, vCampos, vCriterio, "Matrícula", False, "")

  If vHandle > 0 Then
    CurrentQuery.FieldByName("MATRICULA").AsInteger = vHandle
    CurrentQuery.FieldByName("BENEFICIARIO").Clear 
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State = 3 Then
    BOTAOALTERARSETORCARGO.Enabled = False
  Else
    BOTAOALTERARSETORCARGO.Enabled = True
  End If	
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Sql As Object, PRONTUARIO As Object
  Dim IDADE2 As Integer
  Dim Anos As Long, Meses As Long, Dias As Long


  If CurrentQuery.FieldByName("ALTURA").IsNull Then
    MsgBox "Campo Altura deve ser preenchido"
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("PESO").IsNull Then
    MsgBox "Campo Peso deve ser preenchido"
    CanContinue = False
    Exit Sub
  End If

  Set Sql = NewQuery

  If (CurrentQuery.FieldByName("TIPOPACIENTE").AsInteger <> 3) Then
    If Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
      Sql.Active = False
      Sql.Clear
      Sql.Add("SELECT MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO")
      Sql.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
      Sql.Active = True

      CurrentQuery.FieldByName("MATRICULA").AsInteger = Sql.FieldByName("MATRICULA").AsInteger
    End If
  End If

  Set PRONTUARIO = CreateBennerObject("CLIPRONTUARIO.Rotinas")

  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT  DATANASCIMENTO FROM SAM_MATRICULA WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
  Sql.Active = True

  If Not Sql.EOF Then
    PRONTUARIO.Idade(CurrentSystem, Sql.FieldByName("DATANASCIMENTO").AsDateTime, Dias, Meses, Anos)
    IDADE2 = Anos
    CurrentQuery.FieldByName("IDADE").Value = IDADE2
  End If

  Set Sql = Nothing
End Sub

Public Sub VERIFICAREXAMES_OnClick()
  'If CurrentQuery.FieldByName("TIPOFUNCIONARIO").AsInteger <>2 Then
  '  Set obj =CreateBennerObject("MS.FuncionarioExames")
  '  obj.Exec
  '  Set obj =Nothing
  'Else
  '  MsgBox " Este PACIENTE não é funcionário "
  'End If
End Sub


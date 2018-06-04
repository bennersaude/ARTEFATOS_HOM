'HASH: 69254C32E77A1DF96316031DA66F5B12
  '#Uses "*bsShowMessage"


Public Sub ESPECIALIDADE_OnPopup(ShowPopup As Boolean)
  Dim handlexx As Long
  Dim vPrestadorClinica As Long
  Dim vColunas As String
  Dim vWhere As String
  Dim vClinica As String
  Dim ProcuraDll As Object
  If CurrentQuery.FieldByName("CLINICA").AsInteger = 0 Then
    bsShowMessage("É necessário informar a clínica!", "I")
    Exit Sub
  End If
  ShowPopup = False
  Set ProcuraDll = CreateBennerObject("Procura.Procurar")
  Set busca = NewQuery
  vColunas = "SAM_ESPECIALIDADE.DESCRICAO"
  vClinica = Str(CurrentQuery.FieldByName("CLINICA").AsInteger)

  If CurrentQuery.FieldByName("RECURSO").AsInteger = 0 Then
    vWhere = "HANDLE IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = (SELECT PRESTADOR FROM CLI_CLINICA WHERE HANDLE = " + vClinica + "))"
  Else
    vWhere = "HANDLE IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = (SELECT PRESTADOR FROM CLI_RECURSO WHERE HANDLE = " + Str(CurrentQuery.FieldByName("RECURSO").AsInteger) + ")) AND HANDLE IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = (SELECT PRESTADOR FROM CLI_CLINICA WHERE HANDLE = " + vClinica + "))"
  End If

  handlexx = ProcuraDll.Exec(CurrentSystem, "SAM_ESPECIALIDADE", vColunas, 1, "Nome", vWhere, "Especialidades", True, "")
  If handlexx <>0 Then
    busca.Clear
    busca.Add("SELECT PRESTADOR FROM CLI_CLINICA WHERE HANDLE = " + vClinica)
    busca.Active = True
    vPrestadorClinica = busca.FieldByName("PRESTADOR").AsInteger

    busca.Clear
    busca.Add("SELECT HANDLE FROM SAM_PRESTADOR_ESPECIALIDADE")
    busca.Add("WHERE PRESTADOR = :PRESTADOR")
    busca.Add("AND ESPECIALIDADE = :ESPECIALIDADE")
    busca.ParamByName("ESPECIALIDADE").AsInteger = handlexx
    busca.ParamByName("PRESTADOR").AsInteger = vPrestadorClinica
    busca.Active = True
    CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger = busca.FieldByName("HANDLE").AsInteger
  End If
End Sub

Public Sub PROGRAMA_OnPopup(ShowPopup As Boolean)
  Dim handlexx As Long
  Dim vColunas As String
  Dim ProcuraDll As Object
  ShowPopup = False
  Set ProcuraDll = CreateBennerObject("Procura.Procurar")

  vColunas = "CLI_PROGRAMA.DESCRICAO"
  handlexx = ProcuraDll.Exec(CurrentSystem, "CLI_PROGRAMA", vColunas, 1, "Nome", "", "Programas", True, "")
  If handlexx <>0 Then
    CurrentQuery.FieldByName("PROGRAMA").AsInteger = handlexx
  End If
End Sub

Public Sub RECURSO_OnPopup(ShowPopup As Boolean)
  Dim handlexx As Long
  Dim vColunas As String
  Dim vWhere As String
  Dim vClinica As String
  Dim vPrestEspecialidade As String
  Dim ProcuraDll As Object
  Dim busca As Object

  If CurrentQuery.FieldByName("CLINICA").AsInteger = 0 Then
    bsShowMessage("É necessário informar a clínica!", "I")
    Exit Sub
  End If

  ShowPopup = False
  Set ProcuraDll = CreateBennerObject("Procura.Procurar")
  Set busca = NewQuery

  vColunas = "SAM_PRESTADOR.NOME"
  vPrestEspecialidade = Str(CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger)
  vClinica = Str(CurrentQuery.FieldByName("CLINICA").AsInteger)

  If CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger <>0 Then
    vWhere = "SAM_PRESTADOR.HANDLE IN (SELECT PRESTADOR FROM SAM_PRESTADOR_ESPECIALIDADE WHERE ESPECIALIDADE IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE HANDLE = " + vPrestEspecialidade + " )) AND SAM_PRESTADOR.HANDLE IN (SELECT PRESTADOR FROM CLI_RECURSO WHERE CLINICA = " + vClinica + ")"
  Else
    vWhere = "SAM_PRESTADOR.HANDLE IN (SELECT PRESTADOR FROM CLI_RECURSO WHERE CLINICA = " + vClinica + ")"
  End If

  handlexx = ProcuraDll.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, "Nome", vWhere, "Prestadores", True, "")
  If handlexx <>0 Then
    busca.Clear
    busca.Add("SELECT HANDLE FROM CLI_RECURSO WHERE CLINICA = :CLINICA AND PRESTADOR = :PRESTADOR")
    busca.ParamByName("PRESTADOR").AsInteger = handlexx
    busca.ParamByName("CLINICA").AsInteger = CInt(vClinica)
    busca.Active = True
    CurrentQuery.FieldByName("RECURSO").AsInteger = busca.FieldByName("HANDLE").AsInteger
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.RECURSO FROM CLI_ATENDIMENTO A WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("AND EXISTS(SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R WHERE R.HANDLE = A.RECURSO AND R.PRESTADOR = RU.PRESTADOR AND RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("CLINICA").IsNull Then
    bsShowMessage("O campo clínica é obrigatório!", "E")
    CanContinue = False
  End If
  If(Not CurrentQuery.FieldByName("CLINICA").IsNull)Then
  If(CurrentQuery.FieldByName("ESPECIALIDADE").IsNull)And(CurrentQuery.FieldByName("RECURSO").IsNull)And(CurrentQuery.FieldByName("PROGRAMA").IsNull)Then
  bsShowMessage("É necessário inserir pelo menos um dos três itens restantes!", "E")
  CanContinue = False
End If
End If
End Sub


Public Sub CLINICA_OnPoPup(ShowPopup As Boolean)
  Dim handlexx As Long
  Dim vColunas As String
  Dim vWhere As String
  Dim ProcuraDll As Object
  Dim busca As Object
  ShowPopup = False
  Set ProcuraDll = CreateBennerObject("Procura.Procurar")
  Set busca = NewQuery

  vColunas = "SAM_PRESTADOR.NOME"
  vWhere = "SAM_PRESTADOR.HANDLE IN (SELECT PRESTADOR FROM CLI_CLINICA)"
  handlexx = ProcuraDll.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, "Nome", vWhere, "Clínicas", True, "")

  If handlexx <>0 Then
    busca.Clear
    busca.Add("SELECT HANDLE FROM CLI_CLINICA WHERE PRESTADOR = :PRESTADOR")
    busca.ParamByName("PRESTADOR").AsInteger = handlexx
    busca.Active = True
    CurrentQuery.FieldByName("CLINICA").AsInteger = busca.FieldByName("HANDLE").AsInteger
    CurrentQuery.FieldByName("RECURSO").Value = Null
    CurrentQuery.FieldByName("ESPECIALIDADE").Value = Null
  End If
  Set ProcuraDll = Nothing
End Sub


'HASH: 70DD917ABA1ED1D9B70EEE6E4FFF215A
'Tabela: AT_ATEND


'#Uses "*ProcuraEvento"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "Z_NOME|BENEFICIARIO"

  vCriterio = ""
  vCampos = "Nome|Código"

  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCampos, vCriterio, "Beneficiários", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub DATAINCLUSAO_OnExit()
  If Not CurrentQuery.FieldByName("DATAINCLUSAO").IsNull Then
    CurrentQuery.FieldByName("ANO").Value = CurrentQuery.FieldByName("DATAINCLUSAO").Value
  End If
End Sub

Public Sub ENDERECOATENDIMENTO_OnExit()
  MOSTRAENDERECO
End Sub

Public Sub ENDERECOATENDIMENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT PRESTADOR FROM AT_CLINICA WHERE HANDLE = :CLINICA")
  SQL.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  SQL.Active = True

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "LOGRADOURO|NUMERO"

  vCriterio = "PRESTADOR = " + SQL.FieldByName("PRESTADOR").AsString
  vCampos = "Endereço|Número"

  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR_ENDERECO", vColunas, 1, vCampos, vCriterio, "Endereços", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ENDERECOATENDIMENTO").Value = vHandle
  End If
  Set interface = Nothing
End Sub


Public Sub EVENTOPRINCIPAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOPRINCIPAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOPRINCIPAL").Value = vHandle
  End If
End Sub


Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vTabela = "SAM_GRAU"
  vColunas = "SAM_GRAU.DESCRICAO|SAM_GRAU.XTHM|SAM_GRAU.VERIFICAGRAUSVALIDOS"
  vCriterio = "SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' " + _
              "OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=" + _
              CurrentQuery.FieldByName("EVENTOPRINCIPAL").AsString + _
              " AND GE.GRAU=SAM_GRAU.HANDLE))"
  vCampos = "Descrição|XTHM|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Graus válidos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim CLI As Object
  Dim HANDLECLI As Long
  Set CLI = NewQuery
  'PROCURA O HANDLE Do PRESTADOR QUE É A CLÍNICA
  CLI.Add("SELECT PRESTADOR                ")
  CLI.Add("  FROM AT_CLINICA               ")
  CLI.Add(" WHERE HANDLE = :HANDLECLINICA    ")
  CLI.ParamByName("HANDLECLINICA").Value = CurrentQuery.FieldByName("CLINICA").Value
  CLI.Active = True
  HANDLECLI = CLI.FieldByName("PRESTADOR").AsInteger
  Set CLI = Nothing

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabelas As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vTabelas = "SAM_PRESTADOR|SAM_PRESTADOR_PRESTADORDAENTID[SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PRESTADORDAENTID.PRESTADOR]"
  vColunas = "SAM_PRESTADOR.NOME|SAM_PRESTADOR.PRESTADOR" 'CAMPOS DA TABELA
  vCriterio = "ENTIDADE = " + Str(HANDLECLI)
  vCampos = "NOME|PRESTADOR" 'TÍTULO DOS CAMPOS

  vHandle = interface.Exec(CurrentSystem, vTabelas, vColunas, 1, vCampos, vCriterio, "Prestadores", True, PRESTADOR.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub PROGRAMA_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("AT_PROGRAMA")
End Sub

Public Sub TABLE_AfterInsert()
  Dim vSequencia As Long
  NewCounter("AT_ATEND", CurrentQuery.FieldByName("HANDLE").AsInteger, 1, vSequencia)
  CurrentQuery.FieldByName("SEQUENCIA").AsInteger = vSequencia

  CurrentQuery.FieldByName("TABTIPOATENDIMENTO").AsInteger = NodeInternalCode
End Sub

Public Sub TABLE_AfterPost()

  Dim queryBenef As Object
  Dim qInsert As Object
  Dim qDelete As Object
  Set queryBenef = NewQuery
  Set qInsert = NewQuery
  Set qDelete = NewQuery

  Dim VERIFICAEVENTO As Object
  Set VERIFICAEVENTO = NewQuery
  Dim EVENTO As Object
  Set EVENTO = NewQuery


  If CurrentQuery.FieldByName("TABTIPOATENDIMENTO").Value = 2 Then
    'Verifica se o evento já existe
    VERIFICAEVENTO.Clear
    VERIFICAEVENTO.Add("SELECT 1 FROM AT_ATEND_REALIZADOCOL")
    VERIFICAEVENTO.Add(" WHERE EVENTO = :EVENTO")
    VERIFICAEVENTO.Add("   AND ATENDIMENTO = :ATENDIMENTO")
    VERIFICAEVENTO.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOPRINCIPAL").AsInteger
    VERIFICAEVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    VERIFICAEVENTO.Active = True

    If VERIFICAEVENTO.EOF Then

	  If Not InTransaction Then StartTransaction

      EVENTO.Add("INSERT INTO AT_ATEND_REALIZADOCOL  ")
      EVENTO.Add("       (HANDLE,                    ")
      EVENTO.Add("        ATENDIMENTO,               ")
      EVENTO.Add("        EVENTO,                    ")
      EVENTO.Add("        GRAU)                      ")
      EVENTO.Add(" VALUES                            ")
      EVENTO.Add("       (:HANDLE,                   ")
      EVENTO.Add("        :ATENDIMENTO,              ")
      EVENTO.Add("        :EVENTO,                   ")
      EVENTO.Add("        :GRAU)                     ")
      EVENTO.ParamByName("EVENTO").DataType = ftInteger
      EVENTO.ParamByName("GRAU").DataType = ftInteger
      EVENTO.ParamByName("HANDLE").Value = NewHandle("AT_ATEND_REALIZADOCOL")
      EVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      EVENTO.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOPRINCIPAL").Value
      EVENTO.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").Value
      EVENTO.ExecSQL

      If InTransaction Then Commit
    End If

    'Deletar os participantes
    qDelete.Active = False

    If Not InTransaction Then StartTransaction

    qDelete.Add("DELETE FROM AT_ATEND_BENEF WHERE ATENDIMENTO=:ATENDIMENTO")
    qDelete.ParamByName("Atendimento").Value = CurrentQuery.FieldByName("handle").Value
    qDelete.ExecSQL

    If InTransaction Then Commit

    queryBenef.Active = False
    queryBenef.Add("SELECT A.BENEFICIARIO ")
    queryBenef.Add("  FROM AT_TURMA_PARTICIPANTE A, ")
    queryBenef.Add("       AT_TURMA B")
    queryBenef.Add(" WHERE B.HANDLE = :HANDLE")
    queryBenef.Add("   AND A.TURMA  = B.HANDLE")
    queryBenef.ParamByName("Handle").Value = CurrentQuery.FieldByName("turma").Value
    queryBenef.Active = True

    If Not queryBenef.EOF Then

	  If Not InTransaction Then StartTransaction

      qInsert.Active = False
      qInsert.Add("INSERT INTO AT_ATEND_BENEF")
      qInsert.Add(" (HANDLE,BENEFICIARIO,INSCRICAODATA,INSCRICAOUSUARIO,PARTICIPOU,ATENDIMENTO)")
      qInsert.Add(" VALUES (:HANDLE,:BENEFICIARIO,:INSCRICAODATA,:INSCRICAOUSUARIO,:PARTICIPOU,:ATENDIMENTO)")

      While Not queryBenef.EOF 'Inserir os participantes de acordo com os beneficiários cadastrados na TURMA
        qInsert.ParamByName("HANDLE").Value = NewHandle("AT_ATEND_BENEF")
        qInsert.ParamByName("BENEFICIARIO").Value = queryBenef.FieldByName("BENEFICIARIO").Value
        qInsert.ParamByName("INSCRICAODATA").Value = Date
        qInsert.ParamByName("INSCRICAOUSUARIO").Value = CurrentUser
        qInsert.ParamByName("PARTICIPOU").Value = "N"
        qInsert.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("handle").Value
        qInsert.ExecSQL

        queryBenef.Next

		If InTransaction Then Commit
      Wend
    End If
  Else
    'Verifica se o evento já existe
    VERIFICAEVENTO.Clear
    VERIFICAEVENTO.Add("SELECT 1 FROM AT_ATEND_REALIZADOCOL")
    VERIFICAEVENTO.Add(" WHERE EVENTO = :EVENTO")
    VERIFICAEVENTO.Add("   AND ATENDIMENTO = :ATENDIMENTO")
    VERIFICAEVENTO.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOPRINCIPAL").AsInteger
    VERIFICAEVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    VERIFICAEVENTO.Active = True

    If VERIFICAEVENTO.EOF Then

	  If Not InTransaction Then StartTransaction

      EVENTO.Add("INSERT INTO AT_ATEND_REALIZADOMED  ")
      EVENTO.Add("       (HANDLE,                    ")
      EVENTO.Add("        ATENDIMENTO,               ")
      EVENTO.Add("        EVENTO,                    ")
      EVENTO.Add("        GRAU,                      ")
      EVENTO.Add("        QUANTIDADE)                ")
      EVENTO.Add(" VALUES                            ")
      EVENTO.Add("       (:HANDLE,                   ")
      EVENTO.Add("        :ATENDIMENTO,              ")
      EVENTO.Add("        :EVENTO,                   ")
      EVENTO.Add("        :GRAU,                     ")
      EVENTO.Add("        :QUANTIDADE)               ")
      EVENTO.ParamByName("EVENTO").DataType = ftInteger
      EVENTO.ParamByName("GRAU").DataType = ftInteger
      EVENTO.ParamByName("HANDLE").Value = NewHandle("AT_ATEND_REALIZADOMED")
      EVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      EVENTO.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOPRINCIPAL").AsInteger
      EVENTO.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
      EVENTO.ParamByName("QUANTIDADE").Value = 1
      EVENTO.ExecSQL

	  If InTransaction Then Commit
    End If
  End If

  Set qInsert = Nothing
  Set queryBenef = Nothing
  Set qDelete = Nothing
  Set EVENTO = Nothing
  Set VERIFICAEVENTO = Nothing

  
End Sub

Public Sub TABLE_AfterScroll()
  MOSTRAENDERECO
End Sub


Public Function MOSTRAENDERECO As Boolean
  If CurrentQuery.FieldByName("TABTIPOATENDIMENTO").AsInteger = 1 Then
    Dim ENDERECO As Object
    Set ENDERECO = NewQuery
    ENDERECO.Clear
    ENDERECO.Add("SELECT P.NUMERO, M.NOME MUNICIPIO, E.NOME ESTADO FROM SAM_PRESTADOR_ENDERECO P, MUNICIPIOS M, ESTADOS E")
    ENDERECO.Add("WHERE P.MUNICIPIO = M.HANDLE AND P.ESTADO = E.HANDLE AND P.HANDLE = :ENDERECO")
    ENDERECO.ParamByName("ENDERECO").Value = CurrentQuery.FieldByName("ENDERECOATENDIMENTO").AsInteger
    ENDERECO.Active = True

    ROTULONUMERO.Text = ENDERECO.FieldByName("NUMERO").AsString
    ROTULOMUNICIPIO.Text = ENDERECO.FieldByName("MUNICIPIO").AsString + "    " + ENDERECO.FieldByName("ESTADO").AsString
    Set ENDERECO = Nothing
  End If

End Function

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("PROCESSADO").Value = "S" Then
    MsgBox("Operação Inválida!Atendimento já foi processado!")
    CanContinue = False
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim MyQuery As Object

  If CurrentQuery.FieldByName("TABTIPOATENDIMENTO").Value = 2 Then
    Set MyQuery = NewQuery
    MyQuery.Active = False
    MyQuery.Add("SELECT DATAINICIAL,DATAFINAL FROM AT_TURMA WHERE HANDLE=:pHANDLE")
    MyQuery.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("TURMA").Value
    MyQuery.Active = True
    If(Not MyQuery.FieldByName("DATAFINAL").IsNull)Then
    If(CurrentQuery.FieldByName("DATAINCLUSAO").Value <MyQuery.FieldByName("DATAINICIAL").Value)Or(CurrentQuery.FieldByName("DATAINCLUSAO").Value >MyQuery.FieldByName("DATAFINAL").Value)Then
    MsgBox("Data de Inclusão fora da data de vigência da turma informada!")
    Set MyQuery = Nothing
    CanContinue = False
  End If
  ElseIf(CurrentQuery.FieldByName("DATAINCLUSAO").Value <MyQuery.FieldByName("DATAINICIAL").Value)Then
  MsgBox("Data de Inclusão deverá ser maior ou igual a data inicial da vigência da turma informada!")
  Set MyQuery = Nothing
  CanContinue = False
End If
End If

End Sub

Public Sub TABTIPOATENDIMENTO_OnChanging(AllowChange As Boolean)
  If TABTIPOATENDIMENTO.PageIndex <>NodeInternalCode Then
    AllowChange = False
  End If
End Sub


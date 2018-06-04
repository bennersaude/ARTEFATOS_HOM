'HASH: 30397D0B4D8DFAA8D74681103B2FD46E
   Public Sub ODONTOGRAMA_OnClick()
    Dim ACESSO As String
    Dim AGENDA As Long
    Dim ODONTO As Object
    Dim PERMISSAO As Object
    Dim SITUACAO As Object
    Set PERMISSAO =NewQuery
    PERMISSAO.Clear
    PERMISSAO.Add("Select A.RECURSO")
    PERMISSAO.Add("  FROM CLI_ATENDIMENTO A,")
    PERMISSAO.Add("       CLI_RECURSO R")
    PERMISSAO.Add(" WHERE A.HANDLE = :ATENDIMENTO")
    PERMISSAO.Add("   And A.RECURSO = R.HANDLE")
    PERMISSAO.Add("   And EXISTS(Select 1")
    PERMISSAO.Add("                FROM CLI_RECURSO_USUARIO RU")
    PERMISSAO.Add("               WHERE RU.PRESTADOR = R.PRESTADOR")
    PERMISSAO.Add("                 And RU.USUARIO = :USUARIO)")
    PERMISSAO.ParamByName("ATENDIMENTO").AsInteger =RecordHandleOfTable("CLI_ATENDIMENTO")
    PERMISSAO.ParamByName("USUARIO").AsInteger =CurrentUser
    PERMISSAO.Active =True
    If PERMISSAO.EOF Then
      MsgBox("Usuário inválido!")
      Exit Sub
    End If
    If Not PERMISSAO.EOF Then
      PERMISSAO.Clear
      PERMISSAO.Add("SELECT PRONTUARIOODONTOLOGICO FROM CLI_RECURSO_USUARIO USUARIO WHERE USUARIO = :USUARIO")
      PERMISSAO.ParamByName("USUARIO").AsInteger =CurrentUser
      PERMISSAO.Active =True
      ACESSO =PERMISSAO.FieldByName("PRONTUARIOODONTOLOGICO").AsString
      If ACESSO <>"S" Then
        MsgBox("Usuário sem permissão para acessar prontuário odontológico!")
        Exit Sub
      End If
    End If
    Set SITUACAO =NewQuery
    SITUACAO.Clear
    SITUACAO.Add("Select DATAFINAL, DATACANCELAMENTO FROM CLI_ATENDIMENTO ")
    SITUACAO.Add("WHERE HANDLE = :ATENDIMENTO ")
    SITUACAO.ParamByName("ATENDIMENTO").AsInteger =RecordHandleOfTable("CLI_ATENDIMENTO")
    SITUACAO.Active =True
    If(Not SITUACAO.FieldByName("DATAFINAL").IsNull)Or(Not SITUACAO.FieldByName("DATACANCELAMENTO").IsNull)Then
      MsgBox("O atendimento já foi encerrado!")
      Exit Sub
    End If
    AGENDA =RecordHandleOfTable("CLI_AGENDA")
    Set ODONTO =CreateBennerObject("BSCli006.Rotinas")
    ODONTO.Odontograma(CurrentSystem,AGENDA,0,0,1,0)
    Set ODONTO =Nothing
  End Sub

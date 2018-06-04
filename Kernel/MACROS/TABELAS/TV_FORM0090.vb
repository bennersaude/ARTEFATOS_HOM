'HASH: 28D110A2123697DC87CDE0FC60701E67
 'MACRO TV_FORM0090
 '#Uses "*RecordHandleOfTableInterfacePEG"
 '#Uses "*RefreshNodesWithTableInterfacePEG"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Início")

  '#Uses "*bsShowMessage"
  '#Uses "*LimpaEspaco"

  On Error GoTo Erro

    If Not CurrentQuery.FieldByName("NFNUMERO").IsNull Then
      CurrentQuery.FieldByName("NFNUMERO").AsString = LimpaEspaco(CurrentQuery.FieldByName("NFNUMERO").AsString)
    End If

    Dim qDadosPeg As Object
    Set qDadosPeg = NewQuery

    Dim qUpdatePeg As Object
    Set qUpdatePeg = NewQuery

    Dim qAlteracoesPeg As Object
    Set qAlteracoesPeg = NewQuery

    qDadosPeg.Clear
    qDadosPeg.Add(" SELECT NFNUMERO, DATAEMISSAONOTA, RECIBO ")
    qDadosPeg.Add("   FROM SAM_PEG                   ")
    qDadosPeg.Add("  WHERE HANDLE = :HANDLE          ")
    qDadosPeg.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qDadosPeg.Active = True

    If Not InTransaction Then
      StartTransaction
    End If

    qAlteracoesPeg.Clear
    qAlteracoesPeg.Add(" INSERT INTO SAM_PEG_ALTERACOES         ")
    qAlteracoesPeg.Add("             (HANDLE,                   ")
    qAlteracoesPeg.Add("             TABNOTAFISCALRECIBO,       ")
    If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
      qAlteracoesPeg.Add("             NFNUMERONOVO,              ")
      qAlteracoesPeg.Add("             DATAEMISSAONOTANOVO,       ")

      If Not qDadosPeg.FieldByName("NFNUMERO").AsString = "" Then
        qAlteracoesPeg.Add("           NFNUMEROANTERIOR,          ")
      End If

      If Not qDadosPeg.FieldByName("DATAEMISSAONOTA").IsNull Then
        qAlteracoesPeg.Add("           DATAEMISSAONOTAANTERIOR,   ")
      End If
    Else
      qAlteracoesPeg.Add("             RECIBONOVO,              ")

      If Not qDadosPeg.FieldByName("RECIBO").IsNull Then
        qAlteracoesPeg.Add("           RECIBOANTERIOR,          ")
      End If
    End If

    qAlteracoesPeg.Add("             PEG,                       ")
    qAlteracoesPeg.Add("             TABTIPO,                   ")
    qAlteracoesPeg.Add("             USUARIO,                   ")
    qAlteracoesPeg.Add("             DATA,                      ")

    If Not CurrentQuery.FieldByName("MOTIVO").IsNull Then
      qAlteracoesPeg.Add("           MOTIVO,                    ")
    End If

    qAlteracoesPeg.Add("             OBSERVACOES)               ")

    qAlteracoesPeg.Add("     VALUES (:HANDLE,                   ")
    If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
      qAlteracoesPeg.Add("             1,       ")
      qAlteracoesPeg.Add("             :NUMERONOVO,               ")
      qAlteracoesPeg.Add("             :DATAEMISSAONOTANOVO,      ")

      If Not qDadosPeg.FieldByName("NFNUMERO").AsString = "" Then
        qAlteracoesPeg.Add("           :NUMEROANTERIOR,           ")
      End If

      If Not qDadosPeg.FieldByName("DATAEMISSAONOTA").IsNull Then
        qAlteracoesPeg.Add("           :DATAANTERIOR,             ")
      End If
    Else
      qAlteracoesPeg.Add("             2,       ")
      qAlteracoesPeg.Add("             :RECIBONOVO,               ")

      If Not qDadosPeg.FieldByName("RECIBO").IsNull Then
        qAlteracoesPeg.Add("           :RECIBOANTERIOR,           ")
      End If
    End If

    qAlteracoesPeg.Add("             :PEG,                      ")
    qAlteracoesPeg.Add("             :TAB,                      ")
    qAlteracoesPeg.Add("             :USUARIO,                  ")
    qAlteracoesPeg.Add("             :DATA,                     ")

    If Not CurrentQuery.FieldByName("MOTIVO").IsNull Then
      qAlteracoesPeg.Add("           :MOTIVO,                   ")
    End If

    qAlteracoesPeg.Add("             :OBSERVACOES)              ")

    qAlteracoesPeg.ParamByName("HANDLE").AsInteger = NewHandle("SAM_PEG_ALTERACOES")
    SessionVar("HANDLEPEGALTERACOES") = qAlteracoesPeg.ParamByName("HANDLE").AsString
    qAlteracoesPeg.ParamByName("PEG").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qAlteracoesPeg.ParamByName("TAB").AsInteger = "5"
    qAlteracoesPeg.ParamByName("USUARIO").AsInteger = CurrentUser
    qAlteracoesPeg.ParamByName("DATA").AsDateTime = ServerNow

    If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
      qAlteracoesPeg.ParamByName("NUMERONOVO").AsString = CurrentQuery.FieldByName("NFNUMERO").AsString
      qAlteracoesPeg.ParamByName("DATAEMISSAONOTANOVO").AsDateTime = CurrentQuery.FieldByName("DATAEMISSAONOTA").AsDateTime
      If Not qDadosPeg.FieldByName("NFNUMERO").AsString = "" Then
        qAlteracoesPeg.ParamByName("NUMEROANTERIOR").AsString = qDadosPeg.FieldByName("NFNUMERO").AsString
      End If

      If Not qDadosPeg.FieldByName("DATAEMISSAONOTA").IsNull Then
        qAlteracoesPeg.ParamByName("DATAANTERIOR").AsDateTime = qDadosPeg.FieldByName("DATAEMISSAONOTA").AsDateTime
      End If
    Else
      qAlteracoesPeg.ParamByName("RECIBONOVO").AsInteger = CurrentQuery.FieldByName("RECIBO").AsInteger
      If Not qDadosPeg.FieldByName("RECIBO").IsNull Then
        qAlteracoesPeg.ParamByName("RECIBOANTERIOR").AsInteger = qDadosPeg.FieldByName("RECIBO").AsInteger
      End If
    End If

    If Not CurrentQuery.FieldByName("MOTIVO").IsNull Then
      qAlteracoesPeg.ParamByName("MOTIVO").AsInteger = CurrentQuery.FieldByName("MOTIVO").AsInteger
    End If

    qAlteracoesPeg.ParamByName("OBSERVACOES").AsString = CurrentQuery.FieldByName("OBSERVACOES").AsString
    WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Inserir registro de alterações")
    qAlteracoesPeg.ExecSQL
    WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Registro inserido")

    Set qDadosPeg = Nothing
    Set qAlteracoesPeg = Nothing
    Dim qParametrosIntegracao As BPesquisa
    Set qParametrosIntegracao = NewQuery

    qParametrosIntegracao.Add("SELECT HANDLE")
    qParametrosIntegracao.Add("FROM ADM_PARAMINTEGRACAOCORPBENNER")
    qParametrosIntegracao.Active = True

    qUpdatePeg.Clear
    qUpdatePeg.Add("  UPDATE SAM_PEG                      ")
    If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
      qUpdatePeg.Add("     SET NFNUMERO = :NUMERONOTA,      ")
      qUpdatePeg.Add("         DATAEMISSAONOTA = :DATANOTA  ")
    Else
      qUpdatePeg.Add("     SET RECIBO = :RECIBO      ")
    End If
    If Not qParametrosIntegracao.EOF Then
      qUpdatePeg.Add("        ,IDENTIFICADORPAGAMENTO = :IDENTIFICADORPAGAMENTO      ")
    End If
    qUpdatePeg.Add("   WHERE HANDLE = :HANDLE             ")
    If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
      qUpdatePeg.ParamByName("NUMERONOTA").AsString = CurrentQuery.FieldByName("NFNUMERO").AsString
      qUpdatePeg.ParamByName("DATANOTA").AsDateTime = CurrentQuery.FieldByName("DATAEMISSAONOTA").AsDateTime
    Else
      qUpdatePeg.ParamByName("RECIBO").AsInteger = CurrentQuery.FieldByName("RECIBO").AsInteger
    End If
    qUpdatePeg.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    If Not qParametrosIntegracao.EOF Then
      If CurrentQuery.FieldByName("TABALTERACAO").AsInteger = 1 Then
        qUpdatePeg.ParamByName("IDENTIFICADORPAGAMENTO").AsString = CurrentQuery.FieldByName("NFNUMERO").AsString
      Else
        qUpdatePeg.ParamByName("IDENTIFICADORPAGAMENTO").AsString = CurrentQuery.FieldByName("RECIBO").AsString
      End If
    End If
    WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Atualizar PEG")
    qUpdatePeg.ExecSQL
    WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - PEG atualizado")

    Set qUpdatePeg = Nothing
    Set qParametrosIntegracao = Nothing

    If InTransaction Then
      Commit
    End If

    bsShowMessage("Dados alterados com sucesso!", "I")

    If VisibleMode Then
      RefreshNodesWithTableInterfacePEG("SAM_PEG")
    End If

    WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Fim")

  Exit Sub

    Erro:
      If InTransaction Then
        Rollback
      End If

      WriteBDebugMessage("TV_FORM0090.TABLE_BeforePost - Erro" + Err.Description)

      bsShowMessage("Erro ao alterar os dados da nota fiscal. Tente novamente!", "I")

      Set qDadosPeg = Nothing
      Set qAlteracoesPeg = Nothing
      Set qUpdatePeg = Nothing
End Sub

'HASH: 36CE1EE96BE472D29A7F1FF541E6C292
'Macro: SAM_PRESTADOR_PROC_ESPEC_GRP

'#Uses "*bsShowMessage"

'Mauricio Ibelli -04/01/2002 -sms3165 -Se filial padrao do prestador for nulo não checar responsavel

Option Explicit

Dim Mensagem As String
Dim vUsuario As Boolean

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  SQL.Add("SELECT SAM_PRESTADOR_PROC.PRESTADOR, SAM_PRESTADOR_PROC.DATAFINAL, SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True


  Dim Interface As CSBusinessComponent
  Dim Resultado As String

  Set Interface = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
  Interface.AddParameter(pdtInteger, CLng(RecordHandleOfTable("SAM_PRESTADOR_PROC")))
  Mensagem = CStr(Interface.Execute("ValidarProcesso"))

  Ok = Mensagem = ""

  Set SQL = Nothing
  Set S = Nothing


End Function

Public Sub ESPECIALIDADEGRUPO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim ProcuraGrupo As Long
  Dim SQL, SQL1, SQL2 As Object
  Dim vEspecialidade, vVigencia As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vTabela As String

  Set Interface = CreateBennerObject("Procura.Procurar")
  Set SQL1 = NewQuery
  SQL1.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC WHERE HANDLE = :PRESTADORPROCESSO")
  If VisibleMode Then
  	SQL1.ParamByName("PRESTADORPROCESSO").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_ESPEC")
  Else
    SQL1.ParamByName("PRESTADORPROCESSO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
  SQL1.Active = True


  If SQL1.FieldByName("OPERACAO").Value = 5 Or SQL1.FieldByName("OPERACAO").Value = 6 Then
    vColunas = "SAM_ESPECIALIDADEGRUPO.DESCRICAO|SAM_PRESTADOR_ESPECIALIDADEGRP.DATAINICIAL|SAM_PRESTADOR_ESPECIALIDADEGRP.DATAFINAL"
    vCriterio = "SAM_PRESTADOR_ESPECIALIDADEGRP.PRESTADORESPECIALIDADE = " + SQL1.FieldByName("PRESTADORESPECIALIDADE").AsString
    vCampos = "Grupo de Eventos|Data inicial|Datafinal"
    vTabela = "SAM_PRESTADOR_ESPECIALIDADEGRP|SAM_ESPECIALIDADEGRUPO[SAM_ESPECIALIDADEGRUPO.HANDLE=SAM_PRESTADOR_ESPECIALIDADEGRP.ESPECIALIDADEGRUPO]"
    Set SQL = NewQuery
    SQL.Add("SELECT DESCRICAO FROM SAM_ESPECIALIDADE WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = SQL1.FieldByName("ESPECIALIDADE").Value
    SQL.Active = True
    vEspecialidade = SQL.FieldByName("DESCRICAO").AsString
    SQL.Active = False
    SQL.Clear
    vVigencia = " (Vigência: Data inicial: " + SQL1.FieldByName("DATAINICIAL").AsString + " Data final: " + SQL1.FieldByName("DATAFINAL").AsString + ")"
    ProcuraGrupo = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Grupos de eventos da especialidade " + vEspecialidade + vVigencia, True, "")
    SQL.Active = False
  Else
    vColunas = "DESCRICAO"
    vCriterio = "ESPECIALIDADE = " + SQL1.FieldByName("ESPECIALIDADE").AsString
    vCampos = "Grupo de Eventos"
    vTabela = "SAM_ESPECIALIDADEGRUPO"
    ProcuraGrupo = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Todos os Grupos de Eventos das especialidades", True, "")
  End If


  If SQL1.FieldByName("OPERACAO").Value = 5 Or SQL1.FieldByName("OPERACAO").Value = 6 Then
    SQL.Clear
    SQL.Add("SELECT E.HANDLE, P.DATAINICIAL, P.DATAFINAL                                       ")
    SQL.Add("  FROM SAM_ESPECIALIDADEGRUPO          E                                          ")
    SQL.Add("  JOIN SAM_PRESTADOR_ESPECIALIDADEGRP  P ON (P.ESPECIALIDADEGRUPO = E.HANDLE)     ")
    SQL.Add(" WHERE P.HANDLE = :HANDLE                                                         ")
    SQL.ParamByName("HANDLE").Value = ProcuraGrupo
    '-------------------------------------------------------

    '-------------------------------------------------------
    SQL.Active = False
    SQL.Active = True
    CurrentQuery.FieldByName("DATAINICIAL").Value = SQL.FieldByName("DATAINICIAL").Value
    CurrentQuery.FieldByName("DATAFINAL").Value = SQL.FieldByName("DATAFINAL").Value
    CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value = SQL.FieldByName("HANDLE").Value
    CurrentQuery.FieldByName("PRESTADORESPECIALIDADEGRUPO").Value = ProcuraGrupo
  Else
    If SQL1.FieldByName("OPERACAO").Value = 3 Then
      Set SQL2 = NewQuery
      SQL2.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADEGRP                                                        ")
      SQL2.Add("WHERE PRESTADORESPECIALIDADE = :PRESTADORESPECIALIDADE AND ESPECIALIDADEGRUPO = :ESPECIALIDADEGRUPO ")
      SQL2.ParamByName("PRESTADORESPECIALIDADE").Value = SQL1.FieldByName("PRESTADORESPECIALIDADE").Value
      SQL2.ParamByName("ESPECIALIDADEGRUPO").Value = ProcuraGrupo
      SQL2.Active = True

      If Not SQL2.EOF Then
        If MsgBox("Este grupo ja está cadastrado nesta especialidade." + (Chr(13)) + _
                  "Deseja Alterar seus dados ?", vbYesNo) = vbYes Then
          vColunas = "SAM_ESPECIALIDADEGRUPO.DESCRICAO|SAM_PRESTADOR_ESPECIALIDADEGRP.DATAINICIAL|SAM_PRESTADOR_ESPECIALIDADEGRP.DATAFINAL"
          vCriterio = "SAM_PRESTADOR_ESPECIALIDADEGRP.PRESTADORESPECIALIDADE = " + SQL1.FieldByName("PRESTADORESPECIALIDADE").AsString
          vCampos = "Grupo de Eventos|Data inicial|Datafinal"
          vTabela = "SAM_PRESTADOR_ESPECIALIDADEGRP|SAM_ESPECIALIDADEGRUPO[SAM_ESPECIALIDADEGRUPO.HANDLE=SAM_PRESTADOR_ESPECIALIDADEGRP.ESPECIALIDADEGRUPO]"
          Set SQL = NewQuery
          SQL.Add("SELECT DESCRICAO FROM SAM_ESPECIALIDADE WHERE HANDLE = :HANDLE")
          SQL.ParamByName("HANDLE").Value = SQL1.FieldByName("ESPECIALIDADE").Value
          SQL.Active = True
          vEspecialidade = SQL.FieldByName("DESCRICAO").AsString
          SQL.Active = False
          SQL.Clear
          vVigencia = " (Vigência: Data inicial: " + SQL1.FieldByName("DATAINICIAL").AsString + " Data final: " + SQL1.FieldByName("DATAFINAL").AsString + ")"
          ProcuraGrupo = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Grupos de eventos da especialidade " + vEspecialidade + vVigencia, True, "")
          SQL.Active = False
          SQL.Clear
          SQL.Add("SELECT E.HANDLE, P.DATAINICIAL, P.DATAFINAL                                       ")
          SQL.Add("  FROM SAM_ESPECIALIDADEGRUPO          E                                          ")
          SQL.Add("  JOIN SAM_PRESTADOR_ESPECIALIDADEGRP  P ON (P.ESPECIALIDADEGRUPO = E.HANDLE)     ")
          SQL.Add(" WHERE P.HANDLE = :HANDLE                                                         ")
          SQL.ParamByName("HANDLE").Value = ProcuraGrupo
          '-------------------------------------------------------
          
          '-------------------------------------------------------
          SQL.Active = False
          SQL.Active = True
          CurrentQuery.FieldByName("DATAINICIAL").Value = SQL.FieldByName("DATAINICIAL").Value
          CurrentQuery.FieldByName("DATAFINAL").Value = SQL.FieldByName("DATAFINAL").Value
          CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value = SQL.FieldByName("HANDLE").Value
          CurrentQuery.FieldByName("PRESTADORESPECIALIDADEGRUPO").Value = ProcuraGrupo
          DATAINICIAL.ReadOnly = True
          DATAFINAL.ReadOnly = True
          DATAINICIAL1.ReadOnly = False
          DATAFINAL1.ReadOnly = False
        Else
          If MsgBox("Deseja criar um Novo Registro para este grupo ?", vbYesNo) = vbYes Then
            CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value = ProcuraGrupo
          End If

        End If
      Else
        CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value = ProcuraGrupo
      End If

      Set SQL2 = Nothing
    Else
      CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value = ProcuraGrupo
    End If

  End If


  Set Interface = Nothing
  Set SQL = Nothing
  Set SQL1 = Nothing
  ShowPopup = False

End Sub

Public Sub TABLE_AfterInsert()
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    Exit Sub
  End If

  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_ESPEC_GRP"
  End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  DATAINICIAL1.ReadOnly = True
  DATAFINAL1.ReadOnly = True
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC WHERE HANDLE = :PRESTADORPROCESSO")
  SQL.ParamByName("PRESTADORPROCESSO").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_ESPEC")
  SQL.Active = True
  If SQL.FieldByName("OPERACAO").Value = 6 Then
    DATAINICIAL1.ReadOnly = False
    DATAFINAL1.ReadOnly = False
  End If
  If SQL.FieldByName("OPERACAO").Value = 5 Or SQL.FieldByName("OPERACAO").Value = 6 Then
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  Else
    DATAINICIAL.ReadOnly = False
    DATAFINAL.ReadOnly = False
  End If

  If Not CurrentQuery.FieldByName("PRESTADORESPECIALIDADEGRUPO").IsNull Then
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC WHERE HANDLE = :PRESTADORPROCESSO")
  SQL.ParamByName("PRESTADORPROCESSO").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_ESPEC")
  SQL.Active = True
  If SQL.FieldByName("OPERACAO").Value = 2 Then
    CanContinue = False
    bsShowMessage("O tipo de operação não permite inserir registros nesta carga !", "E")
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim VDataInicialEsp, vDataFinalEsp As Date
  Dim VDataInicialGrp, vDataFinalGrp As Date
  Dim vDataStr As String
  Dim condicao As String
  Dim linha As String


  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC WHERE HANDLE = :PRESTADORPROCESSO")
  SQL.ParamByName("PRESTADORPROCESSO").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").Value
  SQL.Active = True

  '---vigencias no cadastro de grupos da especialidade no prestador ------------------
  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
  condicao = "AND PRESTADOR     = " + SQL.FieldByName("PRESTADOR").AsString
  condicao = condicao + " AND ESPECIALIDADE =  " + SQL.FieldByName("ESPECIALIDADE").AsString
  If Not SQL.FieldByName("PRESTADORESPECIALIDADE").IsNull Then
    condicao = condicao + " AND PRESTADORESPECIALIDADE =  " + SQL.FieldByName("PRESTADORESPECIALIDADE").AsString
  End If
  '----------------------------------------
  If CurrentQuery.FieldByName("PRESTADORESPECIALIDADEGRUPO").IsNull Then
    Dim qGRP, qGRPPROC As Object
    Dim Condicao1, Linha1 As String
    If SQL.FieldByName("OPERACAO").Value <>1 Then
      Set qGRP = NewQuery
      qGRP.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADEGRP WHERE PRESTADORESPECIALIDADE = :PRESTADORESPECIALIDADE")
      If SQL.FieldByName("PRESTADORESPECIALIDADE").AsInteger > 0 Then
        qGRP.ParamByName("PRESTADORESPECIALIDADE").AsInteger = SQL.FieldByName("PRESTADORESPECIALIDADE").AsInteger
      Else
      	qGRP.ParamByName("PRESTADORESPECIALIDADE").DataType = ftInteger
      	qGRP.ParamByName("PRESTADORESPECIALIDADE").Clear
      End If
      qGRP.Active = True
      While Not qGRP.EOF
        Condicao1 = condicao + " AND HANDLE = " + qGRP.FieldByName("HANDLE").AsString
        linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADEGRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADEGRUPO", Condicao1)
        qGRP.Next
      Wend
      Set qGRP = Nothing
    End If
    If linha = "" Then
      Condicao1 = " AND PRESTADORPROCESSO = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
      Linha1 = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_ESPEC_GRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADEGRUPO", Condicao1)
      If Linha1 <>"" Then
        CanContinue = False
        bsShowMessage("Existe outro registro neste processo com vigência intercalada!", "E")
        Exit Sub
      End If
    Else
      CanContinue = False
      bsShowMessage(linha, "E")
      Exit Sub
    End If
    '----------------------------------------
  Else
    If Not CurrentQuery.FieldByName("DATAINICIAL1").IsNull Then
      condicao = condicao + " AND HANDLE <>  " + CurrentQuery.FieldByName("PRESTADORESPECIALIDADEGRUPO").AsString
      linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADEGRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime, CurrentQuery.FieldByName("DATAFINAL1").AsDateTime, "ESPECIALIDADEGRUPO", condicao)
      If linha <>"" Then
        CanContinue = False
        bsShowMessage(linha, "E")
        Exit Sub
      End If
    End If
  End If
  Set INTERFACE = Nothing
  '---vigencias no processo --------------------------------------------------------
  If CurrentQuery.FieldByName("DATAINICIAL1").IsNull Then
    Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
    condicao = "AND PRESTADORPROCESSO     = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
    condicao = condicao + " AND DATAINICIAL1 IS NULL"
    linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_ESPEC_GRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADEGRUPO", condicao)
    If linha <>"" Then
      CanContinue = False
      bsShowMessage("Já existe um registro neste processo com vigência intercalando!", "E")
      Exit Sub
    Else
      Set INTERFACE = Nothing
      Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
      condicao = "AND PRESTADORPROCESSO     = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
      condicao = condicao + " AND DATAINICIAL1 IS NOT NULL"
      linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_ESPEC_GRP", "DATAINICIAL1", "DATAFINAL1", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADEGRUPO", condicao)
      If linha <>"" Then
        CanContinue = False
        bsShowMessage("Já existe um registro neste processo com a nova vigência intercalando!", "E")
        Exit Sub
      End If
    End If
    Set INTERFACE = Nothing
  Else
    Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
    condicao = "AND PRESTADORPROCESSO     = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
    condicao = condicao + " AND DATAINICIAL1 IS NULL"
    linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_ESPEC_GRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime, CurrentQuery.FieldByName("DATAFINAL1").AsDateTime, "ESPECIALIDADEGRUPO", condicao)
    If linha <>"" Then
      CanContinue = False
      bsShowMessage("Já existe um registro neste processo com vigência intercalando!", "E")
      Exit Sub
    Else
      Set INTERFACE = Nothing
      Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
      condicao = "AND PRESTADORPROCESSO     = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
      condicao = condicao + " AND DATAINICIAL1 IS NOT NULL"
      linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_ESPEC_GRP", "DATAINICIAL1", "DATAFINAL1", CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime, CurrentQuery.FieldByName("DATAFINAL1").AsDateTime, "ESPECIALIDADEGRUPO", condicao)
      If linha <>"" Then
        CanContinue = False
        bsShowMessage("Já existe um registro neste processo com a nova vigência intercalando!", "E")
        Exit Sub
      End If
    End If
    Set INTERFACE = Nothing
  End If
  '-----------------------------------------------------------
  ' RestricoesDeVigencias-------------------------------------
  vDataStr = ""
  If SQL.FieldByName("OPERACAO").Value = 1 Then
    If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data inicial não pode ser menor que a data inicial da especialidade", "E")
      Exit Sub
    End If

    If Not SQL.FieldByName("DATAFINAL").IsNull Then
      If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
        CanContinue = False
        bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser nula.", "E")
        Exit Sub
      Else
        If CurrentQuery.FieldByName("DATAFINAL").AsDateTime >SQL.FieldByName("DATAFINAL").AsDateTime Then
          CanContinue = False
          bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser maior que a data final da especialidade", "E")
          Exit Sub
        End If
      End If
    End If

  End If

  If SQL.FieldByName("OPERACAO").Value <>2 And SQL.FieldByName("OPERACAO").Value <>5 Then
    '---
    If(SQL.FieldByName("DATAINICIAL1").IsNull)And(CurrentQuery.FieldByName("DATAINICIAL1").IsNull)Then
    If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data inicial não pode ser menor que a data inicial da especialidade", "E")
      Exit Sub
    End If

    If Not SQL.FieldByName("DATAFINAL").IsNull Then
      If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
        CanContinue = False
        bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser nula.", "E")
        Exit Sub
      Else
        If CurrentQuery.FieldByName("DATAFINAL").AsDateTime >SQL.FieldByName("DATAFINAL").AsDateTime Then
          CanContinue = False
          bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser maior que a data final da especialidade", "E")
          Exit Sub
        End If
      End If
    End If
  End If
  '---
  If(SQL.FieldByName("DATAINICIAL1").IsNull)And(Not CurrentQuery.FieldByName("DATAINICIAL1").IsNull)Then
  If CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime <SQL.FieldByName("DATAINICIAL").AsDateTime Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data inicial da nova vigência não pode ser menor que a data inicial da especialidade", "E")
    Exit Sub
  End If

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL1").IsNull Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final da nova vigência não pode ser nula.", "E")
      Exit Sub
    Else
      If CurrentQuery.FieldByName("DATAFINAL1").AsDateTime >SQL.FieldByName("DATAFINAL").AsDateTime Then
        CanContinue = False
        bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final da nova vigência não pode ser maior que a data final da especialidade", "E")
        Exit Sub
      End If
    End If
  End If
End If
'---
If(Not SQL.FieldByName("DATAINICIAL1").IsNull)And(CurrentQuery.FieldByName("DATAINICIAL1").IsNull)Then
If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAINICIAL1").AsDateTime Then
  CanContinue = False
  bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data inicial não pode ser menor que a data inicial da especialidade", "E")
  Exit Sub
End If

If Not SQL.FieldByName("DATAFINAL1").IsNull Then
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser nula.", "E")
    Exit Sub
  Else
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime >SQL.FieldByName("DATAFINAL1").AsDateTime Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final não pode ser maior que a data final da especialidade", "E")
      Exit Sub
    End If
  End If
End If
End If
'---
If(Not SQL.FieldByName("DATAINICIAL1").IsNull)And(Not CurrentQuery.FieldByName("DATAINICIAL1").IsNull)Then
If CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime <SQL.FieldByName("DATAINICIAL1").AsDateTime Then
  CanContinue = False
  bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data inicial da nova vigência não pode ser menor que a data inicial da especialidade", "E")
  Exit Sub
End If

If Not SQL.FieldByName("DATAFINAL1").IsNull Then
  If CurrentQuery.FieldByName("DATAFINAL1").IsNull Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final da nova vigência não pode ser nula.", "E")
    Exit Sub
  Else
    If CurrentQuery.FieldByName("DATAFINAL1").AsDateTime >SQL.FieldByName("DATAFINAL1").AsDateTime Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !" + Chr(10) + "Motivo: A data final da nova vigência não pode ser maior que a data final da especialidade", "E")
      Exit Sub
    End If
  End If
End If
End If
'---
End If

'-----------------------------------------------------------
Set SQL = Nothing
Set SQL1 = Nothing
Set SQL2 = Nothing
End Sub

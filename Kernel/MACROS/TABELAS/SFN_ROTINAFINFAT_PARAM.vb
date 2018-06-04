'HASH: 80E377BF8F46AC650F8EDF06D5F1E01F
'Macro: SFN_ROTINAFINFAT_PARAM
'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
 Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTipoFaturamento As String
  Dim vCodigoInterno As Long
  Dim vHandleRotinaFin As Long
  Dim SQL As BPesquisa
  Dim vNumeroColuna As Integer


  If NodeInternalCode = 900 Then
    Set SQL = NewQuery()
    vHandleRotinaFin = RecordHandleOfTable("SFN_ROTINAFIN")
    SQL.Add("SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN")
    SQL.Add("WHERE HANDLE = :ROTINAFIN")
    SQL.ParamByName("ROTINAFIN").Value = vHandleRotinaFin
    SQL.Active = True

    vTipoFaturamento = SQL.FieldByName("TIPOFATURAMENTO").AsString

  Else

    vTipoFaturamento = Str(RecordHandleOfTable("SIS_TIPOFATURAMENTO"))

  End If

  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio ="(DATACANCELAMENTO IS NULL OR DATACANCELAMENTO > " + SQLDate(ServerNow) + ") "
  vCriterio =vCriterio +"AND EMPRESA = " +Str(CurrentCompany)+" AND TIPOFATURAMENTO=" + vTipoFaturamento
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos ="Nº do Contrato|Contratante|Data Adesão"

	If IsNumeric(CONTRATO.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_CONTRATO",vColunas,vNumeroColuna,vCampos,vCriterio,"Contratos",True,CONTRATO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value =vHandle
  End If


  Set interface =Nothing

  Set SQL = Nothing

End Sub



Public Sub CONTRATOINICIAL_OnPopup(ShowPopup As Boolean)
 Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTipoFaturamento As String
  Dim vHandleRotinaFin As Long
  Dim SQL As BPesquisa
  Dim vTeste As Long


  On Error GoTo ERRO
     If CONTRATOINICIAL.Text <> "" Then
        vTeste = CInt(CONTRATOINICIAL.Text)
      End If


If NodeInternalCode = 900 Then
  Set SQL = NewQuery()
  vHandleRotinaFin = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Add("SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :ROTINAFIN")
  SQL.ParamByName("ROTINAFIN").Value = vHandleRotinaFin
  SQL.Active = True

  vTipoFaturamento = SQL.FieldByName("TIPOFATURAMENTO").AsString

Else

  vTipoFaturamento = Str(RecordHandleOfTable("SIS_TIPOFATURAMENTO"))

End If
  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio ="(DATACANCELAMENTO IS NULL OR DATACANCELAMENTO > " + SQLDate(ServerNow) + ") "
  vCriterio =vCriterio +"AND EMPRESA = " +Str(CurrentCompany)+" AND TIPOFATURAMENTO=" + vTipoFaturamento
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos ="Nº do Contrato|Contratante|Data Adesão"



  vHandle =interface.Exec(CurrentSystem,"SAM_CONTRATO",vColunas,1,vCampos,vCriterio,"Contratos",True,CONTRATOINICIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOINICIAL").Value =vHandle
  End If

  Set interface =Nothing
  Set SQL = Nothing

  Exit Sub

  Erro:
     bsShowMessage("Informe um contrato válido!", "E")
     Exit Sub
End Sub

Public Sub CONVENIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SIGLA|DESCRICAO"

  vCriterio = "EMPRESA = " + Str(CurrentCompany)
  vCampos = "Sigla|Descrição"


  vHandle = interface.Exec(CurrentSystem, "SAM_CONVENIO", vColunas,2, vCampos, vCriterio, "Convênios", True, CONVENIO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONVENIO").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub FAMILIAFINAL_OnPopup(ShowPopup As Boolean)
    '  FAMILIAFINAL.LocalWhere ="SAM_FAMILIA.FAMILIA >= " + _
  '      "(Select FAMILIA FROM SAM_FAMILIA WHERE SAM_FAMILIA.HANDLE = " + _
  '      CurrentQuery.FieldByName("FAMILIAINICIAL").AsString +")"


  If CurrentQuery.FieldByName("FAMILIAINICIAL").IsNull Then
    ShowPopup = False
    Exit Sub
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim Query As Object
  Dim vFamilia As String

  If FAMILIAFINAL.Text = "" Then
    vFamilia = "-99"
  Else
    vFamilia = FAMILIAFINAL.Text
  End If

  Set Query = NewQuery()
  Query.Active = False
  Query.Clear
  Query.Add("SELECT COUNT(1) QTDE FROM SAM_FAMILIA WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND FAMILIA = " +vFamilia + " AND FAMILIA >= (Select FAMILIA FROM SAM_FAMILIA WHERE SAM_FAMILIA.HANDLE = "+ CurrentQuery.FieldByName("FAMILIAINICIAL").AsString + ")")
  On Error GoTo caracter
  Query.Active = True

  If Query.FieldByName("QTDE").AsInteger = 0 Then
     ShowPopup = False
     Set interface = CreateBennerObject("Procura.Procurar")

     vColunas = "SAM_FAMILIA.FAMILIA|C.NOME|D.NOME|B.DESCRICAO|SAM_FAMILIA.DATACANCELAMENTO"

     vCriterio = "SAM_FAMILIA.CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString
     vCriterio = vCriterio + " AND SAM_FAMILIA.FAMILIA >= " + _
                 "(Select FAMILIA FROM SAM_FAMILIA WHERE SAM_FAMILIA.HANDLE = " + _
                 CurrentQuery.FieldByName("FAMILIAINICIAL").AsString + ")"
     vCampos = "Família|Titular Responsável|Pessoa Responsável|Lotação|Data de cancelamento"

     vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA|*SAM_CONTRATO_LOTACAO B[SAM_FAMILIA.CONTRATO=B.CONTRATO AND SAM_FAMILIA.LOTACAO=B.HANDLE]|*SAM_BENEFICIARIO C[C.HANDLE=SAM_FAMILIA.TITULARRESPONSAVEL]|*SFN_PESSOA D[SAM_FAMILIA.PESSOARESPONSAVEL=D.HANDLE]", vColunas, 1, vCampos, vCriterio, "Famílias", True, FAMILIAINICIAL.Text)

     If vHandle <>0 Then
       CurrentQuery.Edit
       CurrentQuery.FieldByName("FAMILIAFINAL").Value = vHandle
     End If
  End If
  Set Query = Nothing
  Set interface = Nothing
  Exit Sub
  caracter:
  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_FAMILIA.FAMILIA|C.NOME|D.NOME|B.DESCRICAO|SAM_FAMILIA.DATACANCELAMENTO"

  vCriterio = "SAM_FAMILIA.CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString
  vCriterio = vCriterio + " AND SAM_FAMILIA.FAMILIA >= " + _
                 "(Select FAMILIA FROM SAM_FAMILIA WHERE SAM_FAMILIA.HANDLE = " + _
                 CurrentQuery.FieldByName("FAMILIAINICIAL").AsString + ")"
  vCampos = "Família|Titular Responsável|Pessoa Responsável|Lotação|Data de cancelamento"

  vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA|*SAM_CONTRATO_LOTACAO B[SAM_FAMILIA.CONTRATO=B.CONTRATO AND SAM_FAMILIA.LOTACAO=B.HANDLE]|*SAM_BENEFICIARIO C[C.HANDLE=SAM_FAMILIA.TITULARRESPONSAVEL]|*SFN_PESSOA D[SAM_FAMILIA.PESSOARESPONSAVEL=D.HANDLE]", vColunas, 1, vCampos, vCriterio, "Famílias", True, FAMILIAINICIAL.Text)

  If vHandle <>0 Then
     CurrentQuery.Edit
     CurrentQuery.FieldByName("FAMILIAFINAL").Value = vHandle
  End If
  Set Query = Nothing
  Set interface = Nothing
End Sub

Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vTeste As Long

  '  CONTRATOFINAL.LocalWhere ="SAM_CONTRATO.CONTRATO >= " + _
'      "(Select CONTRATO FROM SAM_CONTRATO WHERE SAM_CONTRATO.HANDLE = " + _
'      CurrentQuery.FieldByName("CONTRATOINICIAL").AsString +")"


  On Error GoTo ERRO
      If CONTRATOFINAL.Text <> "" Then
         vTeste = CInt(CONTRATOFINAL.Text)
      End If

  If CurrentQuery.FieldByName("CONTRATOINICIAL").IsNull Then
     ShowPopup =False
     Exit Sub
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTipoFaturamento As String
  Dim vCodigoInterno As Long
  Dim vHandleRotinaFin As Long
  Dim SQL As BPesquisa


  If NodeInternalCode = 900 Then
    Set SQL = NewQuery()
    vHandleRotinaFin = RecordHandleOfTable("SFN_ROTINAFIN")
    SQL.Add("SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN")
    SQL.Add("WHERE HANDLE = :ROTINAFIN")
    SQL.ParamByName("ROTINAFIN").Value = vHandleRotinaFin
    SQL.Active = True

    vTipoFaturamento = SQL.FieldByName("TIPOFATURAMENTO").AsString

  Else

    vTipoFaturamento = Str(RecordHandleOfTable("SIS_TIPOFATURAMENTO"))

  End If

  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio ="(DATACANCELAMENTO IS NULL OR DATACANCELAMENTO > " + SQLDate(ServerNow) + ") "
  vCriterio =vCriterio +"AND EMPRESA = " +Str(CurrentCompany)+" AND TIPOFATURAMENTO=" + vTipoFaturamento
  vCriterio =vCriterio +" AND SAM_CONTRATO.CONTRATO >= " + _
                          "(Select CONTRATO FROM SAM_CONTRATO WHERE SAM_CONTRATO.HANDLE = " + _
                           CurrentQuery.FieldByName("CONTRATOINICIAL").AsString +")"
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos ="Nº do Contrato|Contratante|Data Adesão"



  vHandle =interface.Exec(CurrentSystem,"SAM_CONTRATO",vColunas,1,vCampos,vCriterio,"Contratos",True,CONTRATOFINAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOFINAL").Value =vHandle
  End If

  Set interface =Nothing
  Set SQL = Nothing

  Exit Sub

  Erro:
     bsShowMessage("Informe um contrato válido!", "I")
     Exit Sub

End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim SQLRotFin As Object
  Dim HandleRotinaFinFat As Long
  Set SQLRotFin = NewQuery
  HandleRotinaFinFat = RecordHandleOfTable("SFN_ROTINAFINFAT")
  SQLRotFin.Add("SELECT A.ROTINAFIN, A.TABTIPOPROCESSO, B.SITUACAO FROM SFN_ROTINAFINFAT A , SFN_ROTINAFIN B")
  SQLRotFin.Add("WHERE A.HANDLE = :ROTINAFINFAT")
  SQLRotFin.Add("AND B.HANDLE = A.ROTINAFIN")
  SQLRotFin.ParamByName("ROTINAFINFAT").Value = HandleRotinaFinFat

  SQLRotFin.Active = True
  If SQLRotFin.FieldByName("TABTIPOPROCESSO").AsInteger <>3 Then ' <>Inscrição
    If SQLRotFin.FieldByName("SITUACAO").Value = "P" Then
      CanContinue = False
      bsShowMessage("A Rotina já foi processada!", "E")
    End If
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Sub

Public Sub FAMILIAINICIAL_OnPopup(ShowPopup As Boolean)
    If CurrentQuery.FieldByName("CONTRATO").IsNull Then
    ShowPopup = False
    Exit Sub
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim Query As Object
  Dim vFamilia As String

  If FAMILIAINICIAL.Text = "" Then
    vFamilia = "-99"
  Else
    vFamilia = FAMILIAINICIAL.Text
  End If
  Set Query = NewQuery()
  Query.Active = False
  Query.Clear
  Query.Add("SELECT COUNT(1) QTDE FROM SAM_FAMILIA WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND FAMILIA = " +vFamilia)
  On Error GoTo caracter
  Query.Active = True
  If Query.FieldByName("QTDE").AsInteger = 0 Then
    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vColunas = "SAM_FAMILIA.FAMILIA|C.NOME|D.NOME|B.DESCRICAO|SAM_FAMILIA.DATACANCELAMENTO"

    vCriterio = "SAM_FAMILIA.CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString
    vCampos = "Família|Titular Responsável|Pessoa Responsável|Lotação|Data de cancelamento"

    vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA|*SAM_CONTRATO_LOTACAO B[SAM_FAMILIA.CONTRATO=B.CONTRATO AND SAM_FAMILIA.LOTACAO=B.HANDLE]|*SAM_BENEFICIARIO C[C.HANDLE=SAM_FAMILIA.TITULARRESPONSAVEL]|*SFN_PESSOA D[SAM_FAMILIA.PESSOARESPONSAVEL=D.HANDLE]", vColunas, 1, vCampos, vCriterio, "Famílias", True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("FAMILIAINICIAL").Value = vHandle
    End If

    Set interface = Nothing
  End If
  Set Query = Nothing
  Exit Sub
  caracter:
  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_FAMILIA.FAMILIA|C.NOME|D.NOME|B.DESCRICAO|SAM_FAMILIA.DATACANCELAMENTO"

  vCriterio = "SAM_FAMILIA.CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString
  vCampos = "Família|Titular Responsável|Pessoa Responsável|Lotação|Data de cancelamento"

  vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA|*SAM_CONTRATO_LOTACAO B[SAM_FAMILIA.CONTRATO=B.CONTRATO AND SAM_FAMILIA.LOTACAO=B.HANDLE]|*SAM_BENEFICIARIO C[C.HANDLE=SAM_FAMILIA.TITULARRESPONSAVEL]|*SFN_PESSOA D[SAM_FAMILIA.PESSOARESPONSAVEL=D.HANDLE]", vColunas, 1, vCampos, vCriterio, "Famílias", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("FAMILIAINICIAL").Value = vHandle
  End If

  Set interface = Nothing

End Sub

Public Sub GRUPOCONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "DESCRICAO"

  vCriterio = "EMPRESA = " + Str(CurrentCompany)
  vCampos = "Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRUPOCONTRATO", vColunas, 1, vCampos, vCriterio, "Convênios", True, GRUPOCONTRATO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRUPOCONTRATO").Value = vHandle
  End If

  Set interface = Nothing

End Sub

Public Sub TABFATURAR_OnChanging(AllowChange As Boolean)
  If NodeInternalCode <> 900 Then
    VerificaSeProcessada(AllowChange)
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If NodeInternalCode <> 900 Then
    VerificaSeProcessada(CanContinue)
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If NodeInternalCode <> 900 Then
    VerificaSeProcessada(CanContinue)
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  If NodeInternalCode <> 900 Then
    VerificaSeProcessada(CanContinue)

    If CanContinue Then
       Dim SQL As BPesquisa
       Set SQL =NewQuery

       SQL.Clear
       SQL.Add("SELECT TABTIPOPROCESSO")
       SQL.Add("FROM SFN_ROTINAFINFAT")
       SQL.Add("WHERE HANDLE = :HROTINA")
       SQL.ParamByName("HROTINA").Value =RecordHandleOfTable("SFN_ROTINAFINFAT")
       SQL.Active =True

       If SQL.FieldByName("TABTIPOPROCESSO").AsInteger <>3 Then 'Inscrição
          If VerificaSeRateio Then
             SQL.Clear
             SQL.Add("SELECT HANDLE")
             SQL.Add("FROM SFN_ROTINAFINFAT_PARAM")
             SQL.Add("WHERE ROTINAFINFAT = :HANDLE")
             SQL.ParamByName("HANDLE").Value =RecordHandleOfTable("SFN_ROTINAFINFAT")
             SQL.Active =True
             If Not SQL.EOF Then
                CanContinue =False
                bsShowMessage("Para Custo Operacional - Rateio só é permitido um registro de seleção!", "E")
             End If
          End If
       End If
       Set SQL =Nothing
    End If

    'If CanContinue =False Then
    '   CanContinue =True
    '   CurrentQuery.Cancel
    '   RefreshNodesWithTable("SFN_ROTINAFINFAT_PARAM")
    'End If
  End If

End Sub

Public Function VerificaSeRateio As Boolean
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO")
  SQL.Add("WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SIS_TIPOFATURAMENTO")
  SQL.Active = True
  If SQL.FieldByName("CODIGO").AsInteger = 140 Then
    VerificaSeRateio = True
  Else
    VerificaSeRateio = False
  End If
  SQL.Active = False
  Set SQL = Nothing
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)
       Dim SQL As BPesquisa
       Set SQL =NewQuery

       SQL.Clear
       SQL.Add("SELECT LOCALFATURAMENTO, TABTIPOPROCESSO")
       SQL.Add("FROM SFN_ROTINAFINFAT")
       SQL.Add("WHERE HANDLE = :HROTINA")
       SQL.ParamByName("HROTINA").Value =RecordHandleOfTable("SFN_ROTINAFINFAT")
       SQL.Active =True

  '     If SQL.FieldByName("TABTIPOPROCESSO").AsInteger <>3 Then 'Inscrição
  '        If VerificaSeRateio Then
  '           If CurrentQuery.FieldByName("TABFATURAR").Value <>1 Then
  '              CanContinue =False
  '              MsgBox("Para Custo Operacional - Rateio só é possível seleção por Grupo de Contrato")
  '           End If
  '        End If
  '     End If
  '     Set SQL =Nothing


   '85757 - Crys

   'Permitir salvar registros cujo tipo seja "Família" (TABFATURAR = 4) apenas nos seguintes casos:
   '1) Local de faturamento da rotina igual a "Família"
   '2) Tipo Do processo da rotina igual a "Inscrição"

        If (CurrentQuery.FieldByName("TABFATURAR").Value = 4) And _
           (SQL.FieldByName("TABTIPOPROCESSO").AsInteger <> 5)Then
            If (SQL.FieldByName("LOCALFATURAMENTO").AsString = "F") Or (SQL.FieldByName("TABTIPOPROCESSO").AsInteger = 3) Then
                CanContinue =True
            Else
                CanContinue =False
                bsShowMessage("Tipo da rotina não permite filtro por família. Filtro por família permitido apenas para rotinas com faturamento na ''Família'' ou rotinas de ''Inscrição'' !", "E")
           End If
        End If
        Set SQL =Nothing

End Sub

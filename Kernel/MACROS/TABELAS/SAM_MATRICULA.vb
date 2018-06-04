'HASH: BB4556B62C52A9F81A23F497523C0B46
'Macro: SAM_MATRICULA
'Celso Lara -14/12/2001 -Integretor
'Mauricio Ibelli -27/11/2000 -para alteracao de nome/data de nascimento dar a possibilidade
'                               de geracao de cartao
'Última alteração: Milton/17/01/2002 -SMS 5976
'Juliana alterado em 30/09/2002 para mostrar a idade.
'#Uses "*bsShowMessage"

Option Explicit

Dim vgAlteracaoCadastral As Boolean
Dim vgNomeAnterior       As String
Dim vgDataNascimento     As Date
Dim vbControlarTransacao As Boolean
Dim vsModoEdicao         As String

Public Sub BOTAOCADASTRODIGITAL_OnClick()
  'SMS 82810 - Rodrigo Andrade - 24/01/2008

  If (CurrentQuery.State = 1) Then
    Dim vsErro As String
    Dim viRetorno As Long
    Dim interface As Object

    Set interface = CreateBennerObject("BSBIO001.BiometriaImpressaoDigital")

    viRetorno = interface.Cadastro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsErro)

    Set interface = Nothing

    If viRetorno <> 0 Then
      bsShowMessage(vsErro, "I")
    End If

    RefreshNodesWithTable("SAM_MATRICULA")

  Else
    bsShowMessage("O Registro Não Pode estar em edição","I")
  End If

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim OLECancelarMatricula As Object
  Dim pHandleMatricula As Long

  pHandleMatricula = CurrentQuery.FieldByName("HANDLE").AsInteger

  Set OLECancelarMatricula = CreateBennerObject("Matricula.CancelarMatricula")
  OLECancelarMatricula.Exec(CurrentSystem, pHandleMatricula)

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  RefreshNodesWithTable("SAM_MATRICULA")

  Set OLECancelarMatricula = Nothing
End Sub

Public Sub BOTAOPESQUISARCNS_OnClick()
  CARTAONACIONALSAUDE.SetFocus

  If CurrentQuery.FieldByName("NOME").AsString <> "" And CurrentQuery.FieldByName("CPF").AsString <> "" Then
    On Error GoTo Excecao
      Dim Cns As String
 	  Dim component As CSBusinessComponent
 	  Set component = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamMatriculaBLL, Benner.Saude.Beneficiarios.Business")
 	  component.ClearParameters
 	  component.AddParameter(pdtString, CurrentQuery.FieldByName("NOME").AsString)
 	  component.AddParameter(pdtString, CurrentQuery.FieldByName("CPF").AsString)
 	  component.AddParameter(pdtString, CurrentQuery.FieldByName("NOMEMAE").AsString)
 	  component.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime)

      Cns =  component.Execute("PesquisarCnsCadSus")

      If Cns <> "" Then
   	    CurrentQuery.FieldByName("CARTAONACIONALSAUDE").AsString = Cns
 	  Else
 	    bsShowMessage("Não foi encontrado nenhum Cartão Nacional de Saúde para os dados informados.","I")
 	  End If

 	  Set component = Nothing

 	  Exit Sub

 	Excecao:
      MsgBox(Err.Description)

 	Set component = Nothing

  Else
    bsShowMessage("Para pesquisar o Cartão Nacional de Saúde, é obrigatório informar no mínimo os campos 'Nome' e 'CPF'.","I")
  End If
End Sub

Public Sub TABLE_AfterCancel()
  BOTAOCADASTRODIGITAL.Visible = True
  BOTAOCANCELAR.Visible        = True

  'A edição/inclusão da matrícula pode ser chamada durante a inclusão de um beneficiário
  'Se isto ocorrer a transação será a do beneficiário e não a da matrícula
  If vbControlarTransacao And _
     InTransaction Then
    Rollback
  End If
  BOTAOPESQUISARCNS.Enabled = False
End Sub

Public Sub TABLE_AfterCommitted()

On Error GoTo erro:
  'tratamento para o integretor: celsolara 14/12/2001 22:48
  If (vsModoEdicao = "A") And _
     (VisibleMode Or _
      WebMode) Then
    ' ++++++++++
    ' Procedimento para emissao de cartao avulso caso houve alteracao de carencia
      If (CurrentQuery.FieldByName("NOME").AsString <>vgNomeAnterior) Or _
         (CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime <>vgDataNascimento) Then

        Dim sql1 As Object
        Set sql1 = NewQuery

        If VisibleMode Then
          sql1.Add("SELECT A.HANDLE FROM SAM_BENEFICIARIO A")
          sql1.Add("WHERE  A.MATRICULA = :MATRICULA")
          sql1.Add("  AND (   A.DATACANCELAMENTO IS NULL")
          sql1.Add("       OR A.DATACANCELAMENTO >= :HOJE)")
          sql1.Add("  AND A.DATABLOQUEIO IS NULL")

          'beneficiario suspenso
          sql1.Add("  AND NOT EXISTS (")
          sql1.Add("                  SELECT HANDLE")
          sql1.Add("                    FROM SAM_BENEFICIARIO_SUSPENSAO")
          sql1.Add("                   WHERE BENEFICIARIO = A.HANDLE")
          sql1.Add("  AND DATAINICIAL <= :HOJE")
          sql1.Add("  AND (DATAFINAL Is Null Or DATAFINAL >= :HOJE))")

          'familia suspensa
          sql1.Add("  AND NOT EXISTS (")
          sql1.Add("                  SELECT HANDLE")
          sql1.Add("                    FROM SAM_FAMILIA_SUSPENSAO")
          sql1.Add("                   WHERE FAMILIA = A.FAMILIA")
          sql1.Add("  AND DATAINICIAL <= :HOJE")
          sql1.Add("  AND (DATAFINAL Is Null Or DATAFINAL >= :HOJE))")

          'Contrato suspenso
          sql1.Add("  AND NOT EXISTS (")
          sql1.Add("                  SELECT HANDLE")
          sql1.Add("                    FROM SAM_CONTRATO_SUSPENSAO")
          sql1.Add("                   WHERE CONTRATO = A.CONTRATO")
          sql1.Add("  AND DATAINICIAL <= :HOJE")
          sql1.Add("  AND (DATAFINAL Is Null Or DATAFINAL >= :HOJE))")


          sql1.ParamByName("HOJE").Value = ServerDate
          sql1.ParamByName("MATRICULA").Value = CurrentQuery.FieldByName("MATRICULA").AsInteger
          sql1.Active = True

          If Not sql1.EOF Then
            If bsShowMessage("Confirma emissão de Cartão?.", "Q") = vbYes Then
              StartTransaction

              While Not sql1.EOF

                Dim Samcartao As Object
                Set Samcartao = CreateBennerObject("SAMROTINACARTAO.BENEFICIARIO")
                Samcartao.processar(CurrentSystem, sql1.FieldByName("HANDLE").AsInteger, 1)
                Set Samcartao = Nothing

                sql1.Next

              Wend

              Commit
            End If
          End If
        End If


    Dim dllEspecifico As Object
    Set dllEspecifico = CreateBennerObject("Especifico.uEspecifico")

        If dllEspecifico.BEN_CriaRotinaRecalculo(CurrentSystem) Then

          If CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime <> vgDataNascimento Then
            sql1.Clear
            sql1.Add("SELECT HANDLE FROM SAM_BENEFICIARIO ")
            sql1.Add("WHERE  MATRICULA = :MATRICULA")
            sql1.Add("  AND (   DATACANCELAMENTO IS NULL")
            sql1.Add("       OR DATACANCELAMENTO >= :HOJE)")
            sql1.ParamByName("HOJE").Value = ServerDate
            sql1.ParamByName("MATRICULA").Value = CurrentQuery.FieldByName("MATRICULA").AsInteger
            sql1.Active = True

            If Not sql1.EOF Then

              Dim vPrimeiraCompetencia As Date
              Dim vUltimaCompetencia As Date
              Dim SQL As Object
              Set SQL = NewQuery

              SQL.Clear
              SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
              SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO B")
              SQL.Add("WHERE B.MATRICULA = :HMATRICULA")
              SQL.Add("  AND A.BENEFICIARIO = B.HANDLE")
              SQL.ParamByName("HMATRICULA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
              SQL.Active = True

              vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
              vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

              If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then

                SQL.Clear
                SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
                SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
                SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, BENEFICIARIO,")
                SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
                SQL.Add("VALUES")
                SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 4,")
                SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :BENEFICIARIO,")
                SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

                StartTransaction

                While Not sql1.EOF
                  SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
                  SQL.ParamByName("DATAROTINA").Value = ServerDate
                  SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
                  SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
                  SQL.ParamByName("BENEFICIARIO").Value = sql1.FieldByName("HANDLE").AsInteger
                  SQL.ParamByName("USUARIO").Value = CurrentUser
                  SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
                  SQL.ParamByName("DESCRICAO").Value = "Alteração na data de nascimento matrícula única " + CurrentQuery.FieldByName("MATRICULA").AsString

                  SQL.ExecSQL

                  sql1.Next
                Wend

                Commit

              End If

              Set SQL = Nothing

            End If

          End If

    End If

        If VisibleMode Then
          CurrentQuery.Active = False
          CurrentQuery.Active = True
        End If
      End If

      Set sql1 = Nothing
      Set dllEspecifico = Nothing

      vgDataNascimento = 0
      vgNomeAnterior = ""
      ' Fim do procedimento de carencia
      ' ++++++++++
  End If 'visible mode

  If Not CurrentQuery.IsVirtual Then
      RefreshNodesWithTable("SAM_MATRICULA")
  End If

  Exit Sub

erro:
  If InTransaction Then
    Rollback
  End If
End Sub

Public Sub TABLE_AfterPost()
  Dim UPD As Object
  Set UPD = NewQuery

  UPD.Add("UPDATE SAM_BENEFICIARIO SET NOME = :NOME, Z_NOME = :Z_NOME WHERE SAM_BENEFICIARIO.MATRICULA = :MATRICULA")
  UPD.ParamByName("MATRICULA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  UPD.ParamByName("NOME").Value = CurrentQuery.FieldByName("NOME").AsString
  UPD.ParamByName("Z_NOME").Value = CurrentQuery.FieldByName("Z_NOME").AsString
  UPD.ExecSQL

  Set UPD = Nothing
  'retirado o código daqui e passado para AfterCommitted, a pedido do larini - sms 59930

  'Tratamento da integração com o Corporativo
  Dim SQLBeneficiarios As BPesquisa
  Set SQLBeneficiarios = NewQuery

  SQLBeneficiarios.Add("SELECT HANDLE")
  SQLBeneficiarios.Add("FROM SAM_BENEFICIARIO")
  SQLBeneficiarios.Add("WHERE MATRICULA = :HMATRICULA")
  SQLBeneficiarios.ParamByName("HMATRICULA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLBeneficiarios.Active = True

  If Not SQLBeneficiarios.EOF Then
    While Not SQLBeneficiarios.EOF
      Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
      Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

      TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, SQLBeneficiarios.FieldByName("HANDLE").AsInteger)
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_BENEFICIARIO")
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

      TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")

      SQLBeneficiarios.Next
    Wend
  End If

  BOTAOPESQUISARCNS.Enabled = False

  'A edição/inclusão da matrícula pode ser chamada durante a inclusão de um beneficiário
  'Se isto ocorrer a transação será a do beneficiário e não a da matrícula
  If VisibleMode And _
     vbControlarTransacao And _
     InTransaction And _
     CurrentQuery.IsVirtual Then
    Commit
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State = 1 Then
    BOTAOCADASTRODIGITAL.Visible = True
    BOTAOCANCELAR.Visible        = True
  Else
    BOTAOCADASTRODIGITAL.Visible = False
    BOTAOCANCELAR.Visible        = False
  End If

  Dim qBeneficiario As Object
  Dim qEmpresa As Object 'André - SMS 39488
  Dim idade_benef As Integer
  Dim vDias, vMeses, vAnos
  Set qEmpresa = NewQuery 'André - SMS 39488

  IDADE.Text = ""

  If CurrentQuery.State <>3 Then
    idade_benef = CalculaIdadeBeneficiario(0, ServerDate)
  End If

  If idade_benef <> -1 Then
    IDADE.Text = "Idade:" + Str(idade_benef)
  End If

  'André - SMS 39488 - 18/04/2005
  qEmpresa.Clear
  qEmpresa.Add (" SELECT E.NOME                                           ")
  qEmpresa.Add ("   FROM SAM_MATRICULA_EMPRESAPACIENTE ME,                ")
  qEmpresa.Add ("        SAM_EMPRESAPACIENTE            E                 ")
  qEmpresa.Add ("  WHERE E.HANDLE = ME.EMPRESAPACIENTE                    ")
  qEmpresa.Add ("    AND ME.DATAINICIAL <= :DATA                          ")
  qEmpresa.Add ("    AND (ME.DATAFINAL IS Null Or ME.DATAFINAL >= :DATA)  ")
  qEmpresa.Add ("    AND (ME.MATRICULA = :MATRICULA)                      ")
  qEmpresa.Add ("    AND (E.SITUACAO = 'A')                               ")
  qEmpresa.ParamByName("DATA").Value = ServerDate
  qEmpresa.ParamByName("MATRICULA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qEmpresa.Active = True

  If Not (qEmpresa.FieldByName("NOME").IsNull) Then
    LBLEMPRESA.Text = "Empresa do paciente: " + qEmpresa.FieldByName("NOME").AsString
  Else
    LBLEMPRESA.Text = ""
  End If


  'FIM SMS 39488

  'Inicio SMS 82810 - Rodrigo Andrade - 24/01/2008
  Dim qParametrosBenef As Object

  Set qParametrosBenef = NewQuery

  qParametrosBenef.Clear
  qParametrosBenef.Add("SELECT TABBIOMETRIAIMPRESSAODIGITAL FROM SAM_PARAMETROSBENEFICIARIO")
  qParametrosBenef.Active = True

  If (qParametrosBenef.FieldByName("TABBIOMETRIAIMPRESSAODIGITAL").AsInteger = 1) Then
    TABCADASTROIMPRESSAODIGITAL.Visible = False
    BOTAOCADASTRODIGITAL.Visible        = False
  Else
    TABCADASTROIMPRESSAODIGITAL.Visible = True
    BOTAOCADASTRODIGITAL.Visible        = True
  End If

  Set qParametrosBenef = Nothing

  If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then
    BOTAOPESQUISARCNS.Enabled = True
  Else
    BOTAOPESQUISARCNS.Enabled = False
  End If

  'Fim SMS 82810

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  vsModoEdicao = "D"
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vsModoEdicao = "A"

  'A edição/inclusão da matrícula pode ser chamada durante a inclusão de um beneficiário
  'Se isto ocorrer a transação será a do beneficiário e não a da matrícula
  If InTransaction Then
    vbControlarTransacao = False
  Else
    vbControlarTransacao = True
  End If

  BOTAOCADASTRODIGITAL.Visible = False
  BOTAOCANCELAR.Visible        = False

  CIDFALECIMENTO.AnyLevel = True

  vgNomeAnterior   = CurrentQuery.FieldByName("NOME").AsString
  vgDataNascimento = CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime

  BOTAOPESQUISARCNS.Enabled = True

End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  vsModoEdicao = "I"

  'A edição/inclusão da matrícula pode ser chamada durante a inclusão de um beneficiário
  'Se isto ocorrer a transação será a do beneficiário e não a da matrícula
  If InTransaction Then
    vbControlarTransacao = False
  Else
    vbControlarTransacao = True
  End If

  BOTAOCADASTRODIGITAL.Visible = False
  BOTAOCANCELAR.Visible        = False

  CIDFALECIMENTO.AnyLevel = True
  BOTAOPESQUISARCNS.Enabled = True

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vsRetorno  As Long
  Dim interface  As Object
  Dim vsErro     As String
  Dim vsPisPasep As Integer
  Dim viRetorno  As Long

  vsPisPasep = Len(CurrentQuery.FieldByName("PISPASEP").AsString)

  If vsPisPasep > 11 Then
    CurrentQuery.FieldByName("PISPASEP").AsString = Mid(CurrentQuery.FieldByName("PISPASEP").AsString, vsPisPasep - 11, 11)
  End If

  vsErro = ""
  Set interface = CreateBennerObject("BSBEN023.Matricula")

  ' as verificações acerda do nome do beneficiário e da mãe do beneficiário precisam ser efetuadas antes do beforePost do delphi,
  '   porque ela não permite salvar a matricula em alguns casos que podem ser criticados pelas verificações de homonimos e critica sib.
  viRetorno = interface.ValidarNome(CurrentSystem, CurrentQuery.TQuery, vsErro)

  If viRetorno = 1 Then
    bsShowMessage(vsErro, "E")

    CanContinue = False
    Set interface = Nothing
    Exit Sub
  End If

  If viRetorno = 2 Then
    If bsShowMessage("Foram encontradas as seguintes restrições: " + Chr(13) + Chr(13) + vsErro + Chr(13) + Chr(13) + "Deseja continuar?", "Q") = vbNo Then
      If (Not WebMode) Then
        CanContinue = False
      End If

      Set interface = Nothing
      Exit Sub
    End If
  End If

  viRetorno = interface.BeforePost(CurrentSystem, CurrentQuery.TQuery, 0, vsErro)
  Set interface = Nothing

  If (viRetorno = 2) And (vsErro <> "") Then
    If bsShowMessage(vsErro, "Q") = vbNo Then
      If (Not WebMode) Then
        CanContinue = False
      End If
      Exit Sub
    End If

  ElseIf (viRetorno = 1) And (vsErro <> "") Then
	If bsShowMessage( vsErro+Chr(13) + "Deseja continuar?", "Q") = vbNo Then
		If (Not WebMode) Then
	    	CanContinue = False
	    	Exit Sub
	  	End If
	  	Set interface = Nothing
	End If
  End If

  If VisibleMode Then
    If CurrentQuery.IsVirtual And Not InTransaction And vbControlarTransacao Then
      StartTransaction
    End If
  End If
End Sub

Function CalculaIdadeBeneficiario(ByVal pBeneficiario As Long, ByVal pDataAtendimento As Date)As Integer
  Dim vDias           As Integer
  Dim vMeses          As Integer
  Dim vAnos           As Integer
  Dim VDataNascimento As Date
  Dim Query           As Object

  Set Query = NewQuery

  If CurrentQuery.FieldByName("DATANASCIMENTO").IsNull Then
   CalculaIdadeBeneficiario = 0
   Exit Function
  End If

  VDataNascimento = CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime
  If (VDataNascimento > ServerDate) Then
    CalculaIdadeBeneficiario = 0
  Else
    DiferencaData2 pDataAtendimento, VDataNascimento, vDias, vMeses, vAnos
    CalculaIdadeBeneficiario = vAnos
  End If
End Function

Public Sub DiferencaData2(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
  Dim DtSwap As Date
  Dim Day1, Day2, Month1, Month2, Year1, Year2 As Integer

  If Data1 >Data2 Then
    DtSwap = Data1
    Data1 = Data2
    Data2 = DtSwap
  End If

  Year1 = Val(Format(Data1, "yyyy"))
  Month1 = Val(Format(Data1, "mm"))
  Day1 = Val(Format(Data1, "dd"))

  Year2 = Val(Format(Data2, "yyyy"))
  Month2 = Val(Format(Data2, "mm"))
  Day2 = Val(Format(Data2, "dd"))

  Anos = Year2 - Year1
  Meses = 0
  Dias = 0
  If Month2 <Month1 Then
    Meses = Meses + 12
    Anos = Anos -1
  End If
  Meses = Meses + (Month2 - Month1)
  If Day2 <Day1 Then
    Dias = Dias + DiasPorMes(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos -1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMes(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMes = 29
    Else
      DiasPorMes = 28
    End If
  End If

End Function

Public Sub TABLE_NewRecord()
  If CurrentBranch > 0 Then
    CurrentQuery.FieldByName("FILIAL").AsInteger = CurrentBranch
  Else
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Add("SELECT FILIALPADRAO")
    qSQL.Add("FROM Z_GRUPOUSUARIOS")
    qSQL.Add("WHERE HANDLE = :HUSUARIO")
    qSQL.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qSQL.Active = True

    If Not qSQL.FieldByName("FILIALPADRAO").IsNull Then
      CurrentQuery.FieldByName("FILIAL").AsInteger = qSQL.FieldByName("FILIALPADRAO").AsInteger
    End If

    Set qSQL = Nothing

  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
   If CommandID = "BOTAOCANCELAR" Then
     BOTAOCANCELAR_OnClick
   End If
End Sub

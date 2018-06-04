'HASH: 255A1EE47996ED6BD95747C2371B1B4E
'SAM_ROTINAIMP_BENEF
'JULIANO 18/10/2000
'JULIANO 04/04/01

'#Uses "*bsShowMessage"


Public Sub BOTAOALTERAR_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString <>"D" Then
   bsShowMessage("O beneficiário deve estar com a situação em duplicado!","I")
  Else
    Dim IMPORTA As Object
    Set IMPORTA = CreateBennerObject("SamImpBenef.Importa")
    IMPORTA.AlteraBeneficiario(CurrentSystem, CurrentQuery.FieldByName("BENEFICIARIO").AsString, _
                               CurrentQuery.FieldByName("HANDLE").AsInteger, _
                               HandleOfTable("SAM_ROTINAIMP_BENEF"))
    Set IMPORTA = Nothing
  End If
End Sub

Public Sub BENEFICIARIOCRIADO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vcolunas, vWhere As String
  Dim vHandlexx As Long
  Dim Q As Object



  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")
  Set Q = NewQuery

  vcolunas = "SAM_MATRICULA.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.DATACANCELAMENTO|SAM_MATRICULA.CPF|SAM_MATRICULA.RG"
  vWhere = ""

  If (Not CurrentQuery.FieldByName("MATRICULAFUNCIONAL").IsNull) Then
    Q.Add("SELECT COUNT(1) NREC FROM SAM_BENEFICIARIO WHERE MATRICULAFUNCIONAL = :MATRICULAFUNCIONAL")
    Q.ParamByName("MATRICULAFUNCIONAL").Value = CurrentQuery.FieldByName("MATRICULAFUNCIONAL").AsString
    Q.Active=True
    If Q.FieldByName("NREC").AsInteger > 0 Then
      vWhere = "MATRICULAFUNCIONAL = '"+CurrentQuery.FieldByName("MATRICULAFUNCIONAL").AsString+"'"
    End If
  End If

  If Not CurrentQuery.FieldByName("CONTRATO").IsNull Then
    If vWhere = "" Then
      vWhere = "CONTRATO = "+CurrentQuery.FieldByName("CONTRATO").AsString
    Else
      vWhere = vWhere + " AND CONTRATO = "+CurrentQuery.FieldByName("CONTRATO").AsString
    End If
  End If

  If vWhere <> "" Then
    vHandlexx = Interface.Exec(CurrentSystem, "SAM_BENEFICIARIO|SAM_MATRICULA[SAM_BENEFICIARIO.MATRICULA = SAM_MATRICULA.HANDLE]", vcolunas, 1, "Nome|Beneficiario|Matrícula Funcional|Código Afinidade|Data Cancelamento|CPF|RG", vWhere, "Procura por Beneficiário", True, "")
  Else
    vHandlexx = Interface.Exec(CurrentSystem, "SAM_BENEFICIARIO|SAM_MATRICULA[SAM_BENEFICIARIO.MATRICULA = SAM_MATRICULA.HANDLE]", vcolunas, 1, "Nome|Beneficiario|Matrícula Funcional|Código Afinidade|Data Cancelamento|CPF|RG", vWhere, "Procura por Beneficiário", False, "")
  End If

  If vHandlexx <> 0 Then
    CurrentQuery.FieldByName("BENEFICIARIOCRIADO").Value = vHandlexx
    Q.Active=False
    Q.Clear
    Q.Add("SELECT CONTRATO, BENEFICIARIO FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
    Q.ParamByName("HANDLE").Value = vHandlexx
    Q.Active=True
    CurrentQuery.FieldByName("BENEFICIARIO").Value = Q.FieldByName("BENEFICIARIO").AsString
    CurrentQuery.FieldByName("CONTRATO").Value = Q.FieldByName("CONTRATO").AsInteger
  End If

  Set Interface = Nothing
  Set Q = Nothing

End Sub




Public Sub BOTAOCONFIRMAMATRICULA_OnClick()
  Dim confirma As Object
  Set confirma = NewQuery

  If Not InTransaction Then StartTransaction

  confirma.Active = False
  confirma.Add("UPDATE SAM_ROTINAIMP_BENEF ")
  confirma.Add("   SET SITUACAO = :SITUACAO")
  confirma.Add(" WHERE HANDLE   = :HANDLE  ")
  confirma.ParamByName("SITUACAO").Value = "P"
  confirma.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  confirma.ExecSQL

  If InTransaction Then Commit

  'Procura a Filial
  Dim VHANDLEROTINA As Long
  Dim VHANDLEFAMILIA As Long
  Dim VHANDLEFILIAL As Long
  Dim VAUTOMATICO As Boolean
  Dim VTIPOLEIAUTE As Integer
  Dim ENCONTRAFILIAL As Object
  Set ENCONTRAFILIAL = NewQuery
  Dim FILIALOK As Object
  Set FILIALOK = NewQuery

  ENCONTRAFILIAL.Active = False
  ENCONTRAFILIAL.Clear
  ENCONTRAFILIAL.Add("Select F.HANDLE HANDLEFAMILIA,                     ")
  ENCONTRAFILIAL.Add("       FL.HANDLE HANDLEFILIAL,                     ")
  ENCONTRAFILIAL.Add("       R.HANDLE HANDLEROTINA,                      ")
  ENCONTRAFILIAL.Add("       R.GERARAUTOMATICO                           ")
  ENCONTRAFILIAL.Add("  FROM SAM_ROTINAIMP_BENEF B,                      ")
  ENCONTRAFILIAL.Add("       SAM_ROTINAIMP_FAM F,                        ")
  ENCONTRAFILIAL.Add("       SAM_ROTINAIMP_FILIAL FL,                    ")
  ENCONTRAFILIAL.Add("       SAM_ROTINAIMP R                             ")
  ENCONTRAFILIAL.Add(" WHERE F.HANDLE  = B.IMPFAM                        ")
  ENCONTRAFILIAL.Add("   And FL.HANDLE = F.ROTINAIMPFILIAL               ")
  ENCONTRAFILIAL.Add("   And R.HANDLE  = FL.ROTINAIMP                    ")
  ENCONTRAFILIAL.Add("   And B.HANDLE  = :IMPORTABENEF                   ")
  ENCONTRAFILIAL.ParamByName("IMPORTABENEF").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  ENCONTRAFILIAL.Active = True

  VHANDLEFAMILIA = ENCONTRAFILIAL.FieldByName("HANDLEFAMILIA").AsInteger
  VHANDLEFILIAL = ENCONTRAFILIAL.FieldByName("HANDLEFILIAL").AsInteger
  VHANDLEROTINA = ENCONTRAFILIAL.FieldByName("HANDLEROTINA").AsInteger
  VAUTOMATICO = ENCONTRAFILIAL.FieldByName("GERARAUTOMATICO").AsBoolean

  'Procura alguma matricula pendente
  Dim VERIFICAPEND As Object
  Set VERIFICAPEND = NewQuery

  VERIFICAPEND.Active = False
  VERIFICAPEND.Clear
  VERIFICAPEND.Add("SELECT A.*")
  VERIFICAPEND.Add("  FROM SAM_ROTINAIMP_FAM A")
  VERIFICAPEND.Add(" WHERE ROTINAIMPFILIAL =:IMPFILIAL")
  VERIFICAPEND.Add("   AND (A.ERRO = 'S'")
  VERIFICAPEND.Add("    OR EXISTS (SELECT B.HANDLE")
  VERIFICAPEND.Add("                 FROM SAM_ROTINAIMP_BENEF B")
  VERIFICAPEND.Add("                WHERE B.IMPFAM = A.HANDLE")
  VERIFICAPEND.Add("                  AND (B.SITUACAO <> 'P' )))")
  VERIFICAPEND.ParamByName("IMPFILIAL").Value = VHANDLEFILIAL

  VERIFICAPEND.Active = True


  If VAUTOMATICO Then
    If VERIFICAPEND.EOF Then

      Dim IMPORTA As Object
      Dim vsRetornoMensagem As Long
      Dim vsMensagemErro As String


      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
        vsRetornoMensagem = IMPORTA.Confirmar(CurrentSystem, VHANDLEROTINA, VHANDLEFILIAL )
      Else
        Set IMPORTA = CreateBennerObject("BSBEN015.ImportarConfirmar")
        vsRetornoMensagem = IMPORTA.Exec(CurrentSystem, VHANDLEROTINA,vsMensagemErro , 0, VHANDLEFILIAL )
      End If

      If vsRetornoMensagem = 1 Then
        bsShowMessage("Ocorreu erro no processo!" + vsMensagemErro ,"I")
        Set IMPORTA = Nothing
        Exit Sub
      End If


      Set IMPORTA = Nothing

      If Not InTransaction Then
        StartTransaction
      End If

      FILIALOK.Clear
      FILIALOK.Add("UPDATE SAM_ROTINAIMP_FILIAL")
      FILIALOK.Add("   Set SITUACAO = :SITUACAO")
      FILIALOK.Add(" WHERE HANDLE   = :HANDLE  ")
      FILIALOK.ParamByName("SITUACAO").Value = "O"
      FILIALOK.ParamByName("HANDLE").Value = VHANDLEFILIAL
      FILIALOK.ExecSQL

      If InTransaction Then
        Commit
      End If

    End If

  Else

    'Caso não ache nenhum registro a situaçao da filial passará para OK
    If VERIFICAPEND.EOF Then

      If Not InTransaction Then StartTransaction

      FILIALOK.Clear
      FILIALOK.Add("UPDATE SAM_ROTINAIMP_FILIAL")
      FILIALOK.Add("   Set SITUACAO = :SITUACAO")
      FILIALOK.Add(" WHERE HANDLE   = :HANDLE  ")
      FILIALOK.ParamByName("SITUACAO").Value = "O"
      FILIALOK.ParamByName("HANDLE").Value = VHANDLEFILIAL
      FILIALOK.ExecSQL

      If InTransaction Then Commit

    End If
  End If

  Set ENCONTRAFILIAL = Nothing
  Set VERIFICAPEND = Nothing
  Set FILIALOK = Nothing
  Set confirma = Nothing
End Sub

Public Sub BOTAOREJEITARMATRICULA_OnClick()
  If CurrentQuery.FieldByName("EHTITULAR").AsString = "S" Then
    Dim SQL As Object
    Dim vsMensagem As String

    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT COUNT(1) QTD")
    SQL.Add("  FROM SAM_ROTINAIMP_BENEF ")
    SQL.Add(" WHERE SITUACAO = 'H'     ")
    SQL.Add("   AND IMPFAM = " + CurrentQuery.FieldByName("IMPFAM").AsString )
    SQL.Add("   AND HANDLE <>" + CurrentQuery.FieldByName("HANDLE").AsString )
    SQL.Active = True

    If SQL.FieldByName("QTD").AsInteger > 0 Then
       vsMensagem = "Para rejeitar a matrícula do titular, é necessário " + Chr(13) + "rejeitar a matrícula dos dependentes!"
       bsShowMessage(vsMensagem, "I")
      Exit Sub
    End If
  End If

  If  CurrentQuery.FieldByName("CONTRATO").IsNull Then
     vsMensagem = "Para rejeitar é preciso informar o contrato."
     bsShowMessage(vsMensagem, "I")
    Exit Sub
  End If

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAO").AsString = "R"
  CurrentQuery.Post
  If WebMode Then
    bsShowMessage("Matrícula rejeitada.", "I")
  End If

  RefreshNodesWithTable("SAM_ROTINAIMP_BENEF")


End Sub

Public Sub CARGO_OnPopup(ShowPopup As Boolean)

  CARGO.LocalWhere = ""

  If Not CurrentQuery.FieldByName("SETOR").IsNull Then
    CARGO.LocalWhere = "CS_CARGOS.CLASSE = " + CurrentQuery.FieldByName("SETOR").AsString
  End If

End Sub

Public Sub SETOR_OnChange()

  CurrentQuery.FieldByName("CARGO").Clear

End Sub

Public Sub SITUACAORH_OnPopup(ShowPopup As Boolean)

  SITUACAORH.LocalWhere = "SAM_SITUACAORH.HANDLE = -1"

  If Not CurrentQuery.FieldByName("CONTRATO").IsNull Then
    If CurrentQuery.FieldByName("MOVIMENTACAO").AsString = "A" Then
      SITUACAORH.LocalWhere = "SAM_SITUACAORH.HANDLE IN (SELECT HANDLE    " + _
                              "                            FROM SAM_SITUACAORH      " + _
                              "                           WHERE CONTRATO = "+CurrentQuery.FieldByName("CONTRATO").AsString+")"
    Else
      SITUACAORH.LocalWhere = "SAM_SITUACAORH.HANDLE IN (SELECT X.HANDLE              " + _
                              "                            FROM SAM_SITUACAORH X      " + _
                              "                           WHERE X.CONTRATO = "+CurrentQuery.FieldByName("CONTRATO").AsString + _
                              "                             AND X.CONTRATO = (SELECT F.CONTRATO          " + _
                              "                                                FROM SAM_ROTINAIMP_FAM F " + _
                              "                                               WHERE F.HANDLE = "+CurrentQuery.FieldByName("IMPFAM").AsString + _
                              "                                             ) " + _
                              "                         )"
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAO").AsString <>"H" Then
    BOTAOCONFIRMAMATRICULA.Enabled = False
  Else
    BOTAOCONFIRMAMATRICULA.Enabled = True

  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString <>"D" Then
    BOTAOALTERAR.Enabled = False
  Else
    BOTAOALTERAR.Enabled = True

  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "R" Or CurrentQuery.FieldByName("SITUACAO").AsString = "O" Then
    BOTAOREJEITARMATRICULA.Enabled = False
  Else
    BOTAOREJEITARMATRICULA.Enabled = True
  End If

  If CurrentQuery.FieldByName("MOVIMENTACAO").AsString = "A" Then
    MOTIVOINCLUSAO.Visible = False
  Else
    MOTIVOINCLUSAO.Visible = True
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    If Not CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull Then
      bsShowMessage("Data de cancelamento é obrigatória quando existir o motivo de cancelamento ","I")
      CanContinue = False
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull Then
      bsShowMessage("Motivo de cancelamento é obrigatório quando existir data de cancelamento ","I")
      CanContinue = False
      Exit Sub
    End If
  End If

  If Not CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT COUNT(1) QTD           ")
    SQL.Add("  FROM SAM_MOTIVOCANCELAMENTO ")
    SQL.Add(" WHERE CODIGO = " + CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsString )
    SQL.Active = True

    If SQL.FieldByName("QTD").AsInteger = 0 Then
      bsShowMessage("Código do motivo de cancelamento não é válido. ","I")
      CanContinue = False
      Exit Sub
    End If
  End If



End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAOCONFIRMAMATRICULA"
      BOTAOCONFIRMAMATRICULA_OnClick
    Case "BOTAOREJEITARMATRICULA"
      BOTAOREJEITARMATRICULA_OnClick
    Case "BOTAOALTERAR"
        BOTAOALTERAR_OnClick
  End Select
End Sub


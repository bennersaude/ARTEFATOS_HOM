'HASH: 7A3B492AA13276878531A094C5632802
'#Uses "*bsShowMessage"

Public Sub CLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIAL").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CONTRATOMOD_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas = "SAM_MODULO.DESCRICAO|SAM_CONTRATO_MOD.DATAADESAO|SAM_CONTRATO_MOD.DATACANCELAMENTO"
  vCampos = "Descrição|Data adesão|Data cancelamento"
  vCriterio = "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO"))

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  vHandle = Procura.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_CONTRATO_MOD.MODULO = SAM_MODULO.HANDLE]", vColunas, 1, vCampos, vCriterio, CONTRATOMOD.Text, True, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOMOD").Value = vHandle
  End If
  Set Procura = Nothing
End Sub

Public Sub CONTRATOMODCOBRANCA_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas = "SAM_MODULO.DESCRICAO|SAM_CONTRATO_MOD.DATAADESAO|SAM_CONTRATO_MOD.DATACANCELAMENTO"
  vCampos = "Descrição|Data adesão|Data cancelamento"
  vCriterio = "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO"))

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  vHandle = Procura.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_CONTRATO_MOD.MODULO = SAM_MODULO.HANDLE]", vColunas, 1, vCampos, vCriterio, CONTRATOMODCOBRANCA.Text, True, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOMODCOBRANCA").Value = vHandle
  End If
  Set Procura = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = True
  Else
    COMPETENCIAFINAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(CurrentQuery.FieldByName("TABMODULO").AsInteger = 2)And _
     (CurrentQuery.FieldByName("CONTRATOMOD").AsInteger <> _
     CurrentQuery.FieldByName("CONTRATOMODCOBRANCA").AsInteger)Then
  CanContinue = False
  bsShowMessage("O módulo do faturamento mínimo e o módulo de cobrança devem ser iguais", "E")
  Exit Sub
End If

If(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
   (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)Then
	bsShowMessage("A Competência final, se informada, deve ser maior ou igual a inicial", "E")
	CanContinue = False
Else
  CanContinue = True
End If
Dim Interface As Object
Dim Linha As String

Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_FATMIN", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "CONTRATO", "")

If Linha = "" Then
  CanContinue = True
Else
  CanContinue = False
  bsShowMessage(Linha, "E")
End If
End Sub


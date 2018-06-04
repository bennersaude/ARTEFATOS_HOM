'HASH: C6CC63E86857AF0C67CA3C0F2010AD7E
'#Uses "*bsShowMessage"

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO|DATACANCELAMENTO"
  vCampos = "Contrato|Contratante|Data adesão|Data cancelamento"
  vCriterio = "CONVENIO = " + Str(CurrentQuery.FieldByName("CONVENIO").AsInteger)

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  vHandle = Procura.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 1, vCampos, vCriterio, CONTRATO.Text, True, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If
  Set Procura = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = False
  Else
    COMPETENCIAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not (CurrentQuery.FieldByName("PARCELAINICIAL").IsNull) And _
          Not (CurrentQuery.FieldByName("PARCELAFINAL").IsNull) Then

    If CurrentQuery.FieldByName("PARCELAFINAL").AsInteger < _
                                CurrentQuery.FieldByName("PARCELAINICIAL").AsInteger Then
      CanContinue = False
      bsShowMessage("A parcela final não pode ser menor que a parcela inicial!", "E")
      Exit Sub
    End If
  End If


  If (Not CurrentQuery.FieldByName("ADESAOFINAL").IsNull) And _
      ( CurrentQuery.FieldByName("ADESAOINICIAL").IsNull) Then
    CanContinue = False
    bsShowMessage("A adesão inicial é obrigatória se a adesão final estiver preenchida!", "E")
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("ADESAOFINAL").IsNull) And _
      (Not CurrentQuery.FieldByName("ADESAOINICIAL").IsNull) Then
    If CurrentQuery.FieldByName("ADESAOFINAL").AsDateTime < _
                                CurrentQuery.FieldByName("ADESAOINICIAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("A adesão final não pode ser menor que a adesão inicial!", "E")
      Exit Sub
    End If
  End If

  Dim Interface As Object
  Dim vLinha As String
  Dim vCriterio As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  vCriterio = " CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString

  vCriterio = vCriterio + " AND CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString

  If Not CurrentQuery.FieldByName("MODULO").IsNull Then
    vCriterio = vCriterio + " AND MODULO = " + CurrentQuery.FieldByName("MODULO").AsString
  End If

  vLinha = Interface.Vigencia(CurrentSystem, "SAM_DESCONTOESCALONADO", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "", vCriterio)

  If vLinha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(vLinha, "E")
  End If
End Sub


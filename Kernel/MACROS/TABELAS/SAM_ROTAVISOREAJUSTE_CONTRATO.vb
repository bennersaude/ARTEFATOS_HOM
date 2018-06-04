'HASH: 6D1DFA083E3A34C0C87C7B57E00C7715
'Macro: SAM_ROTAVISOREAJUSTE_CONTRATO


Public Sub BOTAOPROCESSAR_OnClick()
  Dim MODULO As Object
  Dim BENEFICIARIO As Object
  Dim ROTINA As Object
  Set MODULO = NewQuery
  Set BENEFICIARIO = NewQuery
  Set ROTINA = NewQuery

  ROTINA.Active = False
  ROTINA.Add("SELECT ROT.PROCESSADO")
  ROTINA.Add("  FROM SAM_ROTAVISOREAJUSTE ROT,")
  ROTINA.Add("       SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT")
  ROTINA.Add(" WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  ROTINA.Add("   And ROTCONT.HANDLE = :HANDLECONTRATO ")
  ROTINA.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  ROTINA.Active = True

  If ROTINA.FieldByName("PROCESSADO").AsString = "S" Then
    MsgBox("Essa rotina está processada!" + Chr(13) + "Para processar o contrato, é preciso cancelar a rotina!")
    cancontinue = False
  Else
    Dim interface As Object
    Set interface = CreateBennerObject("SamAvisoReajuste.Geral")
    interface.AvisoContrato(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set interface = Nothing
  End If

  Set MODULO = Nothing
  Set BENEFICIARIO = Nothing
  Set RODITNA = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim MODULO As Object
  Dim BENEFICIARIO As Object
  Dim ROTINA As Object
  Dim FAMILIA As Object
  Set MODULO = NewQuery
  Set BENEFICIARIO = NewQuery
  Set ROTINA = NewQuery
  Set FAMILIA = NewQuery

  ROTINA.Active = False
  ROTINA.Add("SELECT ROT.PROCESSADO")
  ROTINA.Add("  FROM SAM_ROTAVISOREAJUSTE ROT,")
  ROTINA.Add("       SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT")
  ROTINA.Add(" WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  ROTINA.Add("   And ROTCONT.HANDLE = :HANDLECONTRATO ")
  ROTINA.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  ROTINA.Active = True

  If ROTINA.FieldByName("PROCESSADO").AsString = "S" Then
    MsgBox("Essa rotina está processada!" + Chr(13) + "Para excluir o contrato, é preciso cancelar a rotina!")
    CanContinue = False
  Else

    MODULO.Active = False
    MODULO.Clear
    MODULO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_MOD")
    MODULO.Add("WHERE HANDLE IN (SELECT ROTMOD.HANDLE")
    MODULO.Add("                   FROM SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
    MODULO.Add("                        SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
    MODULO.Add("                        SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF,")
    MODULO.Add("                        SAM_ROTAVISOREAJUSTE_MOD ROTMOD")
    MODULO.Add("                  WHERE ROTCONT.HANDLE  = ROTFAM.ROTINAAVISOCONTRATO")
    MODULO.Add("                    And ROTFAM.HANDLE   = ROTBENEF.ROTINAAVISOFAMILIA")
    MODULO.Add("                    And ROTBENEF.HANDLE = ROTMOD.ROTINAAVISOBENEF")
    MODULO.Add("                    And ROTCONT.HANDLE  = :HANDLECONTRATO)")
    MODULO.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    MODULO.ExecSQL

    BENEFICIARIO.Active = False
    BENEFICIARIO.Clear
    BENEFICIARIO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_BENEF")
    BENEFICIARIO.Add("WHERE HANDLE IN (SELECT ROTBENEF.HANDLE")
    BENEFICIARIO.Add("                   FROM SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
    BENEFICIARIO.Add("                        SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
    BENEFICIARIO.Add("                        SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF")
    BENEFICIARIO.Add("                  WHERE ROTCONT.HANDLE  = ROTFAM.ROTINAAVISOCONTRATO")
    BENEFICIARIO.Add("                    And ROTFAM.HANDLE   = ROTBENEF.ROTINAAVISOFAMILIA")
    BENEFICIARIO.Add("                    And ROTCONT.HANDLE  = :HANDLECONTRATO)")
    BENEFICIARIO.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    BENEFICIARIO.ExecSQL

    FAMILIA.Active = False
    FAMILIA.Clear
    FAMILIA.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_FAMILIA")
    FAMILIA.Add("WHERE HANDLE IN (SELECT ROTFAM.HANDLE")
    FAMILIA.Add("                   FROM SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
    FAMILIA.Add("                        SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM")
    FAMILIA.Add("                   WHERE ROTCONT.HANDLE  = ROTFAM.ROTINAAVISOCONTRATO")
    FAMILIA.Add("                     And ROTCONT.HANDLE  = :HANDLECONTRATO)")
    FAMILIA.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    FAMILIA.ExecSQL

  End If

  Set MODULO = Nothing
  Set BENEFICIARIO = Nothing
  Set RODITNA = Nothing
  Set FAMILIA = Nothing
End Sub


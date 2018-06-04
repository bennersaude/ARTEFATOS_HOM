'HASH: 9BD5DB43EFDACE2F85C39D439FC0CD32
'#Uses "*bsShowMessage"

Public Sub BOTAOINSERIRPRESTADOR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_REDEATENDPORTAL_PRESTADOR", "Inserindo Prestadores", "SAM_PRESTADOR", "PRESTADOR", "REDEATENDPORTAL", CurrentQuery.FieldByName("HANDLE").AsInteger, "S", "NOME")
  Set Obj = Nothing

  RefreshNodesWithTable("SAM_REDEATENDPORTAL_PRESTADOR")
End Sub

Public Sub TABLE_AfterScroll()

  If CurrentQuery.State = 3 Then
	BOTAOINSERIRPRESTADOR.Enabled = False
  Else
    BOTAOINSERIRPRESTADOR.Enabled = True
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim QREGISTRODUP As Object
Set QREGISTRODUP = NewQuery

QREGISTRODUP.Active = False
QREGISTRODUP.Clear
QREGISTRODUP.Add("SELECT HANDLE	")
QREGISTRODUP.Add("  FROM SAM_REDEATENDPORTAL ")
QREGISTRODUP.Add(" WHERE REGISTROMS = :REGISTROMS ")
QREGISTRODUP.Add("   AND HANDLE <> :REDEATEND ")
QREGISTRODUP.ParamByName("REGISTROMS").AsInteger = CurrentQuery.FieldByName("REGISTROMS").AsInteger
QREGISTRODUP.ParamByName("REDEATEND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
QREGISTRODUP.Active = True

If Not QREGISTRODUP.FieldByName("HANDLE").IsNull Then
  	bsShowMessage("Plano já cadastrado para esta rede de atendimento do Portal.","I")
  	Set QREGISTRODUP = Nothing
  	CanContinue = False
  	Exit Sub
End If

Set QREGISTRODUP = Nothing

Dim QREGISTROMS As Object

Set QREGISTROMS = NewQuery

QREGISTROMS.Active = False
QREGISTROMS.Clear
QREGISTROMS.Add("SELECT NOVAREGULAMENTACAO FROM SAM_REGISTROMS WHERE HANDLE = :REGISTRO ")
QREGISTROMS.ParamByName("REGISTRO").AsInteger = CurrentQuery.FieldByName("REGISTROMS").AsInteger
QREGISTROMS.Active = True

If QREGISTROMS.FieldByName("NOVAREGULAMENTACAO").AsString = "N" And CurrentQuery.FieldByName("CODIGOPLANOOPERADORA").IsNull Then
  bsShowMessage("Para planos não regulamentados é obrigatório o preenchimento do campo 'Código do plano na operadora'.", "I")
  Set QREGISTROMS = Nothing
  CanContinue = False
  Exit Sub
End If

Set QREGISTROMS = Nothing

End Sub

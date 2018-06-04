'HASH: A5D652B9054A1F42FD93AE6914CA668E
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim StrQ As String
  Dim vTabela As String
  Dim vCampos As String
  Dim vColunas As String



  Set Interface = CreateBennerObject("PROCURA.PROCURAR")

  Dim vPrestador As String
  vPrestador = CurrentQuery.FieldByName("PRESTADOR").AsString
  ShowPopup = False



  StrQ = "EXISTS (SELECT DISTINCT SAM_TGE.HANDLE" + _
           " FROM SAM_PRESTADOR                      P" + _
           "     JOIN SAM_PRESTADOR_ESPECIALIDADE   PE On (PE.PRESTADOR          = P.HANDLE) " + _
           "     JOIN SAM_ESPECIALIDADE             E  On (E.HANDLE              = PE.ESPECIALIDADE) " + _
           "     JOIN SAM_ESPECIALIDADEGRUPO        EG On (E.HANDLE              = EG.ESPECIALIDADE) " + _
           "LEFT JOIN SAM_ESPECIALIDADEGRUPO_EXEC   EE On (EE.ESPECIALIDADEGRUPO = EG.HANDLE)" + _
           "WHERE P.HANDLE = " + vPrestador + _
           "  AND SAM_TGE.HANDLE = EE.EVENTO)"


  vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
  vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"

  CurrentQuery.FieldByName("EVENTO").AsInteger = _
                           Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, StrQ, "", True, "")


End Sub


Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
		If WebMenuCode = "T3880" Then
			PRESTADOR.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
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
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub


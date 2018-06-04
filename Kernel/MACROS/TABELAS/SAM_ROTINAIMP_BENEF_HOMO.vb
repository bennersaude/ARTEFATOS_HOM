'HASH: BD250E71A8049206D95A8EA7B5F776EB
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCONFIRMAMATRICULA_OnClick()
  Dim BSBEN015 As Object
  Dim vsMensagem      As String
  Dim viRetorno As Long

  Set BSBEN015 = CreateBennerObject("BSBEN015.ConfirmarHomonimo")
  viRetorno = BSBEN015.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagem)

  Select Case viRetorno
	Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
    Case 0
        bsShowMessage("Operação Concluída!", "I")
	Case 1
		bsShowMessage(vsMensagem, "I")
  End Select
  Set BSBEN015 = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Dim VerifSituacao As Object
  Set VerifSituacao = NewQuery
  Dim BuscaSexo As Object
  Set BuscaSexo = NewQuery

  VerifSituacao.Active = False
  VerifSituacao.Clear
  VerifSituacao.Add("SELECT  B.SITUACAO SITUACAO             ")
  VerifSituacao.Add("  FROM SAM_ROTINAIMP_BENEF B            ")
  VerifSituacao.Add(" WHERE B.HANDLE  = :IMPORTABENEF        ")
  VerifSituacao.ParamByName("IMPORTABENEF").Value = CurrentQuery.FieldByName("IMPORTABENEF").AsInteger
  VerifSituacao.Active = True

  If VerifSituacao.FieldByName("SITUACAO").AsString <> "H" Then
    BOTAOCONFIRMAMATRICULA.Enabled = False
  Else
    BOTAOCONFIRMAMATRICULA.Enabled = True
  End If

  BuscaSexo.Active = False
  BuscaSexo.Clear
  BuscaSexo.Add("SELECT M.SEXO SEXO                ")
  BuscaSexo.Add("  FROM SAM_MATRICULA M            ")
  BuscaSexo.Add("  WHERE M.HANDLE = :MATRICULAHOMO ")
  BuscaSexo.ParamByName("MATRICULAHOMO").Value = CurrentQuery.FieldByName("MATRICULA").AsInteger
  BuscaSexo.Active = True

  If BuscaSexo.FieldByName("SEXO").AsString = "F" Then
    SEXO.Text = "Sexo: Feminino"
  Else
    SEXO.Text = "Sexo: Masculino"
  End If

  Set VerifSituacao = Nothing
  Set BuscaSexo = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
  Case "BOTAOCONFIRMAMATRICULA"
    BOTAOCONFIRMAMATRICULA_OnClick
  End Select

  CanContinue = True
End Sub

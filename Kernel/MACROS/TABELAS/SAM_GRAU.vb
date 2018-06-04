'HASH: 539C8D55410D85ABD06DF568A67857ED
'Macro: SAM_GRAU
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BITFACE_OnChange() 'CLEBER - ODONTOLÓGICO
  CurrentQuery.FieldByName("VALORFACE").AsInteger = CurrentQuery.FieldByName("BITFACE").AsInteger
End Sub

Public Sub DENTE_OnChange()
	Dim EhUmDente As Boolean

	EhUmDente = EhDente

    BITFACE.Visible = EhUmDente
    VALORFACE.Visible = EhUmDente
End Sub

Public Sub TABLE_AfterScroll() 'CLEBER - ODONTOLÓGICO
  If VisibleMode = False Then
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOGRAU").AsInteger = 2 Then
  	Dim EhUmDente As Boolean
	EhUmDente = EhDente

	BITFACE.Visible = EhUmDente
    VALORFACE.Visible = EhUmDente
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaGrau As Object
  Set qVerificaGrau = NewQuery

  qVerificaGrau.Clear
  qVerificaGrau.Add("SELECT HANDLE FROM SAM_GRAU WHERE GRAU = :GRAU AND HANDLE <> :HANDLE")
  qVerificaGrau.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
  qVerificaGrau.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaGrau.Active = True

  If qVerificaGrau.FieldByName("HANDLE").AsInteger > 0 Then
    CanContinue = False
    bsShowMessage("Grau já existente.", "E")
  End If

Set qVerificaGrau = Nothing


  If VisibleMode = False Then
    Exit Sub
  End If

  Const AUXILIAR = "2"
  If (CurrentQuery.FieldByName("ORIGEMVALOR").AsString = AUXILIAR) Then
    If (CurrentQuery.FieldByName("NUMAUXILIAR").IsNull) Then
      CanContinue = False
      bsShowMessage("Deve inforrmar o número do AUXILIAR", "E")
      NUMAUXILIAR.SetFocus
    End If
  Else
    If (Not CurrentQuery.FieldByName("NUMAUXILIAR").IsNull) And (CurrentQuery.FieldByName("ORIGEMVALOR").AsString <> "4") Then
      CanContinue = False
      bsShowMessage("Informar o número somente para grau auxiliar", "E")
      NUMAUXILIAR.SetFocus
    End If
  End If

  '------  Durval alterado em 03/12/2001 ------------------------------------------------------------------------------------
  '-------- Durval alterado 08/04/2003
  If (CurrentQuery.FieldByName("ORIGEMVALOR").AsInteger <> 1 And (CurrentQuery.FieldByName("PRECOPORGRAU").AsString = "S" Or CurrentQuery.FieldByName("PRECOPORGRAUDOTACAO").AsString = "S")) Then
    bsShowMessage("Preço por grau somente quanto a Origem do valor for Tabela de Dotações", "E")
    CanContinue = False
  End If
  If CurrentQuery.FieldByName("PRECOPORGRAU").AsString = "S" And CurrentQuery.FieldByName("PRECOPORGRAUDOTACAO").AsString = "S" Then
    bsShowMessage("Apenas uma das opcões pode ser escolhida no preço por grau.", "E")
    CanContinue = False
  End If
  '--------------------------------------------------------------------------------------------------------------------------
  '---Cleber - Odontológico
  If CurrentQuery.FieldByName("TABTIPOGRAU").AsInteger = 2 Then
    If Not EhDente Then
      CurrentQuery.FieldByName("BITFACE").AsInteger = 0
      CurrentQuery.FieldByName("VALORFACE").AsInteger = 0
    Else
      CurrentQuery.FieldByName("VALORFACE").AsInteger = CurrentQuery.FieldByName("BITFACE").AsInteger
    End If

    If CurrentQuery.FieldByName("BITFACE").AsInteger = 0 Then
      CurrentQuery.FieldByName("VALORFACE").AsInteger = 0
    End If
  End If
End Sub

Public Function EhDente As Boolean

	EhDente = False

    Dim qCliDente As BPesquisa
    Set qCliDente = NewQuery

    qCliDente.Add("SELECT TIPO FROM CLI_DENTE WHERE HANDLE = :DENTE")
    qCliDente.ParamByName("DENTE").AsInteger = CurrentQuery.FieldByName("DENTE").AsInteger
    qCliDente.Active = True

    If Not qCliDente.EOF Then
		If qCliDente.FieldByName("TIPO").AsString = "D" Then
			EhDente = True
		Else
			If Len(Trim(qCliDente.FieldByName("TIPO").AsString)) = 0 Then
			    bsShowMessage("Tabela de dentes está desatualizada. Processar a verificação '12 - Verificação para atualizar a tabela de sistema com o cadastro de dentes' para visualização correta dos dados odontológicos!", "E")
			End If
		End If
	End If

    Set qCliDente = Nothing
End Function

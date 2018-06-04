'HASH: 9045EED639AAB0E35DDEF1943CA1565F

'MACRO: SFN_MODELO_ESTRUTURA
'27/09/2001 - Milton

'ALTERADO em 17/04/2002
' Milton

'#Uses "*bsShowMessage

Option Explicit

Public Sub Camposvisiveis

  If CurrentQuery.FieldByName("TABTIPO").AsInteger <> 17 Then 'Integraçao com o compras
	If CurrentQuery.FieldByName("TIPO").AsString <> "D" Then
    	CONDICAO.Visible = False
    	QUEBRA.Visible = True
  	Else
    	CAMPOQUEBRA.Visible = False
    	CONDICAO.Visible = True
    	QUEBRA.Visible = False
  	End If

  	If CurrentQuery.FieldByName("QUEBRA").AsString = "N" And CurrentQuery.FieldByName("TIPO").AsString <> "D" Then
    	CAMPOQUEBRA.Visible = True
    	CAMPOQUEBRA.ReadOnly = True
  	Else
    	If CurrentQuery.FieldByName("TIPO").AsString <> "D" Then
      		CAMPOQUEBRA.Visible = True
      		CAMPOQUEBRA.ReadOnly = False
    	End If
  	End If
  Else
    CAMPOQUEBRA.Visible = False
    CONDICAO.Visible = False
    QUEBRA.Visible = False
  End If

End Sub

Public Sub CAMPOQUEBRA_OnPopup(ShowPopup As Boolean)

  Dim Qtipo As BPesquisa
  Set Qtipo = NewQuery
  Dim PModelo As Integer

  PModelo = CurrentQuery.FieldByName("MODELO").AsInteger
  Qtipo.Add("SELECT TABTIPO FROM SFN_MODELO WHERE HANDLE= :PMODELO")
  Qtipo.ParamByName("PMODELO").AsInteger = PModelo
  Qtipo.Active = True

  If Qtipo.FieldByName("TABTIPO").AsInteger = 3 Then
    CAMPOQUEBRA.LocalWhere = "SIS_CONTABCAMPOS.ORIGEM='9'"
  End If
  If Qtipo.FieldByName("TABTIPO").AsInteger = 4 Then
    CAMPOQUEBRA.LocalWhere = "SIS_CONTABCAMPOS.ORIGEM='M'"
  End If
  If (Qtipo.FieldByName("TABTIPO").AsInteger<>3) And (Qtipo.FieldByName("TABTIPO").AsInteger<>4) Then
    CAMPOQUEBRA.LocalWhere = "SIS_CONTABCAMPOS.ORIGEM='4'"
  End If

  Set Qtipo = Nothing

End Sub

Public Sub TABLE_AfterPost()
  	Camposvisiveis

	Dim atualizaModelo As BPesquisa
    Set atualizaModelo = NewQuery

    atualizaModelo.Clear
    atualizaModelo.Add(" UPDATE SFN_MODELO                              ")
    atualizaModelo.Add("    SET USUARIOALTERACAO  = :HUSUARIOALTERACAO, ")
	atualizaModelo.Add("	    DATAHORAALTERACAO = :DATAHORAALTERACAO  ")
    atualizaModelo.Add("  WHERE HANDLE = :HMODELO                       ")
    atualizaModelo.ParamByName("HUSUARIOALTERACAO").AsInteger = CurrentSystem.CurrentUser
    atualizaModelo.ParamByName("DATAHORAALTERACAO").AsDateTime = CurrentSystem.ServerNow
    atualizaModelo.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger
    atualizaModelo.ExecSQL

    Set atualizaModelo = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Camposvisiveis
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Pinheiro - sms 65137
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 15 Then
    If CurrentQuery.FieldByName("TIPODETALHECORPORATIVO").AsInteger <> 1 Then
      If (VisibleMode And DETALHEPAI.Text = "") Or _
      	 ((Not VisibleMode) And CurrentQuery.FieldByName("DETALHEPAI").AsString = "") Then
        bsShowMessage("Detalhe pai deve ser preenchido", "I")
        DETALHEPAI.SetFocus
        CancelDescription = "Detalhe pai deve ser preenchido"
        CanContinue = False
      End If
    End If

    If (CurrentQuery.FieldByName("TIPODETALHECORPORATIVO").AsInteger =1) Then
      If (VisibleMode And DETALHEPAI.Text <> "") Or _
      	 ((Not VisibleMode) And CurrentQuery.FieldByName("DETALHEPAI").AsString <> "") Then
        bsShowMessage("Tipo do detalhe não permite detalhe pai!", "I")
        CancelDescription = "Tipo do detalhe não permite detalhe pai"
        CanContinue = False
      End If
    End If
  End If
  'FIM Pinheiro - sms 65137


  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 17 Then

    If CurrentQuery.FieldByName("TIPO").AsString <> "D" Then
      bsShowMessage("'Tipo' deve ser 'Detalhe' para a opção 'Integração com o compras'!", "I")
      CancelDescription = "'Tipo' deve ser 'Detalhe' para a opção 'Integração com o compras'!"
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("TIPODETALHECORPORATIVO").AsInteger <> 4 Then
      If (VisibleMode And DETALHEPAI.Text = "") Or _
      	 ((Not VisibleMode) And CurrentQuery.FieldByName("DETALHEPAI").AsString = "") Then
        bsShowMessage("Detalhe pai deve ser preenchido", "I")
        DETALHEPAI.SetFocus
        CancelDescription = "Detalhe pai deve ser preenchido"
        CanContinue = False
      End If
    End If

    If (CurrentQuery.FieldByName("TIPODETALHECORPORATIVO").AsInteger =1) Then
      If (VisibleMode And DETALHEPAI.Text <> "") Or _
      	 ((Not VisibleMode) And CurrentQuery.FieldByName("DETALHEPAI").AsString <> "") Then
        bsShowMessage("Tipo do detalhe não permite detalhe pai!", "I")
        CancelDescription = "Tipo do detalhe não permite detalhe pai!"
        CanContinue = False
      End If
    End If


  End If




End Sub

Public Sub TABTIPO_OnChange()

  If TABTIPO.PageIndex = 16 Then
    CAMPOQUEBRA.Visible = False
    CONDICAO.Visible = False
    QUEBRA.Visible = False
  End If

End Sub

Public Sub TIPO_OnChange()
  Camposvisiveis
End Sub

Public Sub DETALHEPAI_OnPopup(ShowPopup As Boolean)
  'Pinheiro - sms 65137
  Dim interface As Object
  Dim vCriterio As String
  Dim vHandle As Long
  Dim vColunas As String
  Dim vCampos As String

  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="DESCRICAO|ORDEM|TABTIPO"

  vCriterio ="MODELO = "+Str(CurrentQuery.FieldByName("MODELO").AsInteger)
  vCriterio =vCriterio +" AND HANDLE <> "+Str(CurrentQuery.FieldByName("HANDLE").AsInteger)
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 15 Then
    vCriterio =vCriterio +" AND TABTIPO = 15" 'Corporativo
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 17 Then
    vCriterio =vCriterio +" AND TABTIPO = 17" 'Corporativo - Integração com o compras
  End If

  vCampos ="Descrição|Ordem|Tipo"

  vHandle =interface.Exec(CurrentSystem,"SFN_MODELO_ESTRUTURA",vColunas,1,vCampos,vCriterio,"Contratos",True,DETALHEPAI.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("DETALHEPAI").Value =vHandle
  End If


  Set interface =Nothing
  'FIM Pinheiro - sms 65137

End Sub

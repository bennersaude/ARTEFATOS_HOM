'HASH: A19494BCA02410432EB4EDFEC08F68D8
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOINSS_OnPopup(ShowPopup As Boolean)

    Dim dllPROCURAProcurar As Object
    Dim vsTabela           As String
    Dim vsCampos           As String
    Dim vsColunas          As String
    Dim vsCriterio         As String
    Dim vsTitulo           As String
    Dim viHandle           As Long

    If (EVENTOINSS.PopupCase <> 0) Then

        ShowPopup = False

        Set dllPROCURAProcurar = CreateBennerObject("PROCURA.Procurar")

        vsTabela   = "SAM_TGE"
        vsCampos   = "ESTRUTURA|DESCRICAO
        vsColunas  = "Estrutura|Descrição
        vsCriterio = "INATIVO = 'N' AND ULTIMONIVEL = 'S'"
        vsTitulo   = "Pesquisa de Evento na TGE para INSS"

        viHandle = dllPROCURAProcurar.Exec(CurrentSystem, vsTabela, vsCampos, 2, vsColunas, vsCriterio, vsTitulo, False, "")

        If (viHandle <> 0) Then
            CurrentQuery.Edit
            CurrentQuery.FieldByName("EVENTOINSS").AsInteger = viHandle
        End If

        Set dllPROCURAProcurar = Nothing

    Else
        ShowPopup = True
    End If

End Sub

Public Sub EVENTOTAXAADM_OnPopup(ShowPopup As Boolean)

    Dim dllPROCURAProcurar As Object
    Dim vsTabela           As String
    Dim vsCampos           As String
    Dim vsColunas          As String
    Dim vsCriterio         As String
    Dim vsTitulo           As String
    Dim viHandle           As Long

    If (EVENTOTAXAADM.PopupCase <> 0) Then

        ShowPopup = False

        Set dllPROCURAProcurar = CreateBennerObject("PROCURA.Procurar")

        vsTabela   = "SAM_TGE"
        vsCampos   = "ESTRUTURA|DESCRICAO
        vsColunas  = "Estrutura|Descrição
        vsCriterio = "INATIVO = 'N' AND ULTIMONIVEL = 'S'"
        vsTitulo   = "Pesquisa de Evento na TGE para Taxa de Administração"

        viHandle = dllPROCURAProcurar.Exec(CurrentSystem, vsTabela, vsCampos, 2, vsColunas, vsCriterio, vsTitulo, False, "")

        If (viHandle <> 0) Then
            CurrentQuery.Edit
            CurrentQuery.FieldByName("EVENTOTAXAADM").AsInteger = viHandle
        End If

        Set dllPROCURAProcurar = Nothing

    Else
        ShowPopup = True
    End If

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim dllProcura_Procurar As Object
  Dim viHandle As Long
  Dim vsCabecs As String
  Dim vsColunas As String
  Dim vsCriterio As String
  Dim vsTabela As String
  Dim vsTitulo As String

  If (PRESTADOR.PopupCase <> 0) Then
    ShowPopup = False
    Set dllProcura_Procurar = CreateBennerObject("Procura.Procurar")

    vsCabecs = "Código|Prestador|CPFCNPJ"
    vsColunas = "PRESTADOR|NOME|CPFCNPJ"
    vsCriterio = "SAM_PRESTADOR.CONVENIORECIPROCIDADE = 'S' "
    vsTabela = "SAM_PRESTADOR"
    vsTitulo = "Prestadores - Convênio de Reciprocidade"

    viHandle = dllProcura_Procurar.Exec(CurrentSystem, vsTabela, vsColunas, 2, vsCabecs, vsCriterio, vsTitulo, False, "")

    If viHandle <> 0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("PRESTADOR").AsInteger = viHandle
    End If
    Set dllProcura_Procurar = Nothing
  Else
    ShowPopup = True
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim dllSamGeral_Vigencia As Object
  Dim vsLinha As String
  Dim vsCondicao As String

  Set dllSamGeral_Vigencia = CreateBennerObject("SAMGERAL.Vigencia")
  vsCondicao = " AND CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " "

  If VisibleMode = True Then
    vsLinha = dllSamGeral_Vigencia.Vigencia(CurrentSystem, "SAM_CONTRATO_CONVRECIP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", vsCondicao)
    If (vsLinha <> "") Then
      CanContinue = False
	  bsShowMessage(vsLinha,"E")
    End If
  End If

  Set dllSamGeral_Vigencia = Nothing
End Sub

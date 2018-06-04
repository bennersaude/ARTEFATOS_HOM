'HASH: 1C32DFE73C861EB16B816DE8C467AE5E
'Macro: SAM_BENEFICIARIO_MOD
'Mauricio -19/10/2000 -Inclusao de codigo tabela de preco para faixa etaria
'Mauricio -29/11/2000 -Dar a opcao de emissao de cartao se incluir/excluir/alterar modulo do Beneficiario
'#Uses "*bsShowMessage"
Option Explicit
Dim vCodigoTabelaPrcAnterior      As String
Dim vsModoEdicao                  As String
Dim vbModuloCanceladoAnterior     As Boolean
Dim vdDataCancelamentoAnterior    As Date
Dim viHMotivoCancelamentoAnterior As Long

Public Sub BOTAOCANCELAR_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long

  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
	  							   1, _
								   "TV_FORM0029", _
								   "Cancelamento de módulo", _
								   0, _
								   180, _
								   420, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
	Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
	Case 1
		bsShowMessage(vsMensagem, "I")
  End Select

  Set BSINTERFACE0002 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOREATIVAR_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long

  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_FORM0034", _
								   "Reativação de módulo", _
								   0, _
								   120, _
								   230, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
	Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
	Case 1
			bsShowMessage(vsMensagem, "I")
  End Select

  Set BSINTERFACE0002 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOTRANSFEREMODULO_OnClick()
  Dim bs As CSBusinessComponent

  Set bs = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.Beneficiarios.TransferenciaModulo, Benner.Saude.Beneficiarios.Business") ' formato: [namespace.classe], [assembly]

  bs.ClearParameters
  bs.AddParameter(pdtInteger, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger)
  If bs.Execute("VerificaSuspensao") Then

    bsShowMessage("Não é permitido transferir o módulo por motivo de suspensão!", "I")
    Exit Sub
  End If

  Dim x As Object
  Set x = CreateBennerObject("BSINTERFACE0018.TransfereModulo")
  x.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set x = Nothing
  RefreshNodesWithTable("SAM_BENEFICIARIO_MOD")
End Sub

Public Sub DATAADESAO_OnChange()
  MODULO.LocalWhere = "DATACANCELAMENTO IS NULL AND DATAADESAO <= " + SQLDate(CurrentQuery.FieldByName("DATAADESAO").AsDateTime)
End Sub

Public Sub MODULO_OnChange()
  If CurrentQuery.State = 3 Then ' Inclusão
    Dim qm As Object
    Set qm = NewQuery

    qm.Clear
    qm.Add("SELECT FM.PRIMEIRAPARCELA, FM.PARCELADIAS, FM.AGENTEAGENCIAVENDAS, FM.TIPOCOMISSAO, FM.SEGUNDAPARCELA")
    qm.Add("  FROM SAM_BENEFICIARIO B, ")
    qm.Add("       SAM_FAMILIA_MOD FM")
    qm.Add(" WHERE B.HANDLE = :BENEFICIARIO")
    qm.Add("   AND FM.MODULO = :CONTRATOMOD")
    qm.Add("   AND FM.FAMILIA = B.FAMILIA")
    qm.Add("   AND FM.DATAADESAO <= :DTADESAO")
    qm.Add("   AND (FM.DATACANCELAMENTO Is Null OR FM.DATACANCELAMENTO >= :DTADESAO)")
    qm.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    qm.ParamByName("CONTRATOMOD").Value = CurrentQuery.FieldByName("MODULO").AsInteger
    qm.ParamByName("DTADESAO").Value = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
    qm.Active = True

    If Not qm.EOF Then
      If(Not qm.FieldByName("PRIMEIRAPARCELA").IsNull)Then
        CurrentQuery.FieldByName("PRIMEIRAPARCELA").Value = qm.FieldByName("PRIMEIRAPARCELA").Value
      End If

      If(Not qm.FieldByName("SEGUNDAPARCELA").IsNull)Then
        CurrentQuery.FieldByName("SEGUNDAPARCELA").Value = qm.FieldByName("SEGUNDAPARCELA").Value
      End If

      If(Not qm.FieldByName("PARCELADIAS").IsNull)Then
        CurrentQuery.FieldByName("PARCELADIAS").Value = qm.FieldByName("PARCELADIAS").Value
      Else
        CurrentQuery.FieldByName("PARCELADIAS").Clear
      End If

      If(Not qm.FieldByName("AGENTEAGENCIAVENDAS").IsNull)Then
        CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Value = qm.FieldByName("AGENTEAGENCIAVENDAS").Value
      Else
        CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Clear
      End If

      If(Not qm.FieldByName("TIPOCOMISSAO").IsNull)Then
        CurrentQuery.FieldByName("TIPOCOMISSAO").Value = qm.FieldByName("TIPOCOMISSAO").Value
      Else
        CurrentQuery.FieldByName("TIPOCOMISSAO").Clear
      End If
    Else
      qm.Clear
      qm.Add("SELECT CM.PRIMEIRAPARCELA, CM.PARCELADIAS, CM.AGENTEAGENCIAVENDAS, CM.TIPOCOMISSAO, CM.SEGUNDAPARCELA")
      qm.Add("FROM SAM_CONTRATO_MOD CM")
      qm.Add("WHERE CM.HANDLE = :CONTRATOMOD")
      qm.ParamByName("CONTRATOMOD").Value = CurrentQuery.FieldByName("MODULO").AsInteger
      qm.Active = True

      If Not qm.EOF Then
        If(Not qm.FieldByName("PRIMEIRAPARCELA").IsNull)Then
          CurrentQuery.FieldByName("PRIMEIRAPARCELA").Value = qm.FieldByName("PRIMEIRAPARCELA").Value
        End If

        If(Not qm.FieldByName("SEGUNDAPARCELA").IsNull)Then
          CurrentQuery.FieldByName("SEGUNDAPARCELA").Value = qm.FieldByName("SEGUNDAPARCELA").Value
        End If

        If(Not qm.FieldByName("PARCELADIAS").IsNull)Then
          CurrentQuery.FieldByName("PARCELADIAS").Value = qm.FieldByName("PARCELADIAS").Value
        Else
          CurrentQuery.FieldByName("PARCELADIAS").Clear
        End If

        If(Not qm.FieldByName("AGENTEAGENCIAVENDAS").IsNull)Then
          CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Value = qm.FieldByName("AGENTEAGENCIAVENDAS").Value
        Else
          CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Clear
        End If

        If(Not qm.FieldByName("TIPOCOMISSAO").IsNull)Then
          CurrentQuery.FieldByName("TIPOCOMISSAO").Value = qm.FieldByName("TIPOCOMISSAO").Value
        Else
          CurrentQuery.FieldByName("TIPOCOMISSAO").Clear
        End If
      End If
    End If

    qm.Active = False
    Set qm = Nothing
  End If
End Sub

Public Sub TABLE_AfterCancel()
  BOTAOTRANSFEREMODULO.Visible  = True
  BOTAOCANCELAR.Visible         = True
  BOTAOREATIVAR.Visible         = True
End Sub

Public Sub TABLE_AfterEdit()
  vCodigoTabelaPrcAnterior = CurrentQuery.FieldByName("CODIGOTABELAPRC").AsString
End Sub

Public Sub TABLE_AfterPost()
  Dim Interface   As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String

  Set Interface = CreateBennerObject("BSBen022.Modulo")

  viResultado = Interface.AfterPost(CurrentSystem, _
                                    CurrentQuery.TQuery, _
                                    vsModoEdicao, _
                                    vCodigoTabelaPrcAnterior, _
                                    vbModuloCanceladoAnterior, _
                                    vsMensagem)

  Set Interface = Nothing

  If viResultado = 1 Then
    Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
  Else
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "I")
    End If
  End If
End Sub

Public Function EhModuloObrigatorio(piHContratoMod As Long) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT OBRIGATORIO                  ")
  SQL.Add("  FROM SAM_CONTRATO_MOD             ")
  SQL.Add(" WHERE HANDLE = :HCONTRATOMOD       ")
  SQL.ParamByName("HCONTRATOMOD").AsInteger = piHContratoMod
  SQL.Active = True

  If SQL.FieldByName("OBRIGATORIO").AsString = "S" Then
    EhModuloObrigatorio = True
  Else
    EhModuloObrigatorio = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub PrepararTransfereModuloWeb()

  Dim SQL As Object
  Set SQL = NewQuery



  SQL.Add("SELECT CONTRATO, NOME")
  SQL.Add("FROM SAM_BENEFICIARIO")
  SQL.Add("WHERE HANDLE = :HBENEFICIARIO")
  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  SessionVar("nomeBeneficiario") = SQL.FieldByName("NOME").AsString

  SQL.Active = False
  SQL.Clear
  Set SQL = NewQuery
  SQL.Add("SELECT PLANO, MODULO FROM SAM_CONTRATO_MOD WHERE HANDLE = :HMODULO")
  SQL.ParamByName("HMODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
  SQL.Active = True


  SessionVar("HANDLEBENEFICIARIO") = CurrentQuery.FieldByName("BENEFICIARIO").AsString
  SessionVar("HANDLEPLANOMODULO") = SQL.FieldByName("PLANO").AsString
  SessionVar("HANDLEMODULO") = SQL.FieldByName("MODULO").AsString


  Set SQL = Nothing

End Sub


Public Sub TABLE_AfterScroll()

  If CurrentQuery.State <> 1 Then
    BOTAOTRANSFEREMODULO.Visible  = False
    BOTAOREATIVAR.Visible         = False
    DATAADESAO.ReadOnly           = False
    MODULO.ReadOnly               = False
  Else
    BOTAOTRANSFEREMODULO.Visible  = True
    BOTAOREATIVAR.Visible         = True
    DATAADESAO.ReadOnly           = True
    MODULO.ReadOnly               = True
  End If

  SessionVar("HMODBENEFICIARIO") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)

  PrepararTransfereModuloWeb

If(Not CurrentQuery.FieldByName("MODULO").IsNull)Then
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT	MODULO.TIPOMODULO, PLANO.DESCRICAO PLANO           				   					           ")
  SQL.Add("  FROM	SAM_CONTRATO                   CONTRATO														   ")
  SQL.Add("  LEFT JOIN SAM_CONTRATO_MOD            CONTRATO_MOD    ON (CONTRATO_MOD.CONTRATO = CONTRATO.HANDLE)    ")
  SQL.Add("  LEFT JOIN SAM_MODULO 		           MODULO	       ON (CONTRATO_MOD.MODULO   = MODULO.HANDLE)      ")
  SQL.Add("  LEFT JOIN SAM_PLANO                   PLANO           ON (CONTRATO_MOD.PLANO    = PLANO.HANDLE)       ")
  SQL.Add(" WHERE CONTRATO_MOD.HANDLE = :PMODULO                                                                   ")


  SQL.ParamByName("PMODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
  SQL.Active = True

  PLANO.Text = "Plano: " + SQL.FieldByName("PLANO").AsString

  'para mostrar o tipo do modulo sms 13506
  If(SQL.FieldByName("TIPOMODULO").IsNull)Then
    TIPOMODULO.Text = "Tipo do Módulo:"
  Else
    If SQL.FieldByName("TIPOMODULO").AsString = "C" Then
      TIPOMODULO.Text = "Tipo do Módulo: Cobertura"
    Else
      If SQL.FieldByName("TIPOMODULO").AsString = "A" Then
        TIPOMODULO.Text = "Tipo do Módulo: Agravo"
      Else
        If SQL.FieldByName("TIPOMODULO").AsString = "G" Then
          TIPOMODULO.Text = "Tipo do Módulo: Garantia Assist. Médica"
        Else
          If SQL.FieldByName("TIPOMODULO").AsString = "O" Then
            TIPOMODULO.Text = "Tipo do Módulo: Outro"
          Else
            If SQL.FieldByName("TIPOMODULO").AsString = "P" Then
              TIPOMODULO.Text = "Tipo do Módulo: Plano Viva Bem"
            Else
              If SQL.FieldByName("TIPOMODULO").AsString = "S" Then
                TIPOMODULO.Text = "Tipo do Módulo: Suplementação de PF"
              End If
            End If
          End If
        End If
      End If
    End If
  End If

  SQL.Active = False
  Set SQL = Nothing
End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  BOTAOTRANSFEREMODULO.Visible  = False
  BOTAOREATIVAR.Visible         = False

  vCodigoTabelaPrcAnterior      = CurrentQuery.FieldByName("CODIGOTABELAPRC").AsString
  vsModoEdicao                  = "A"
  vbModuloCanceladoAnterior     = Not(CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull)
  vdDataCancelamentoAnterior    = CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime
  viHMotivoCancelamentoAnterior = CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  BOTAOTRANSFEREMODULO.Visible  = False
  BOTAOREATIVAR.Visible         = False

  vsModoEdicao = "I"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface   As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String

  Set Interface = CreateBennerObject("BSBen022.Modulo")

  viResultado = Interface.BeforePost(CurrentSystem, _
                                     CurrentQuery.TQuery, _
                                     vbModuloCanceladoAnterior, _
                                     vdDataCancelamentoAnterior, _
                                     viHMotivoCancelamentoAnterior, _
                                     vsMensagem)

  If viResultado = 1 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")
  Else
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "I")
    End If
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_NewRecord()
  MODULO.LocalWhere = "DATACANCELAMENTO IS NULL AND DATAADESAO <= " + SQLDate(ServerDate())
End Sub

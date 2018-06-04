'HASH: CF0B8665B72023CA6D31D162543A1CA4
'TV_FORM0029
'#Uses "*bsShowMessage"
Option Explicit
Dim vsModoEdicao As String

Public Function ModuloCancelado(piHBeneficiarioMod As String) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE                       ")
  SQL.Add("  FROM SAM_BENEFICIARIO_MOD         ")
  SQL.Add(" WHERE HANDLE = :HMODULO            ")
  SQL.Add("   AND DATACANCELAMENTO IS NOT NULL ")
  SQL.ParamByName("HMODULO").AsString = piHBeneficiarioMod
  SQL.Active = True

  ModuloCancelado = (SQL.FieldByName("HANDLE").AsInteger > 0)

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function ModuloObrigatorio(piHBeneficiarioMod As String) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT C.OBRIGATORIO                             ")
  SQL.Add("  FROM SAM_BENEFICIARIO_MOD B                    ")
  SQL.Add("  JOIN SAM_CONTRATO_MOD C ON(C.HANDLE = B.MODULO)")
  SQL.Add(" WHERE B.HANDLE = :HMODULO                       ")
  SQL.ParamByName("HMODULO").AsString = piHBeneficiarioMod
  SQL.Active = True

  ModuloObrigatorio = (SQL.FieldByName("OBRIGATORIO").AsString = "S")

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub TABLE_AfterScroll()
  If WebMode Then
    If InStr(SQLServer,"MSSQL") Then
  	  MODULO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE FROM SAM_CONTRATO_MOD WHERE CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + ") AND DATAADESAO <= CAST(@CAMPO(DATA) AS DATETIME) + 1 AND OBRIGATORIO = 'S' AND HANDLE <> (SELECT MODULO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + " ) )"
    ElseIf InStr(SQLServer,"ORACLE") Then
  	  MODULO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE FROM SAM_CONTRATO_MOD WHERE CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + ") AND DATAADESAO <= TO_DATE(@CAMPO(DATA)) + 1 AND OBRIGATORIO = 'S' AND HANDLE <> (SELECT MODULO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + " ) )"
    End If
  Else
    If InStr(SQLServer,"MSSQL") Then
  	  MODULO.LocalWhere = "SAM_CONTRATO_MOD.HANDLE IN (SELECT HANDLE FROM SAM_CONTRATO_MOD WHERE CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + ") AND DATAADESAO <= CAST(@DATA AS DATETIME) + 1 AND OBRIGATORIO = 'S' AND HANDLE <> (SELECT MODULO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + " ) )"
    ElseIf InStr(SQLServer,"ORACLE") Then
  	  MODULO.LocalWhere = "SAM_CONTRATO_MOD.HANDLE IN (SELECT HANDLE FROM SAM_CONTRATO_MOD WHERE CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + ") AND DATAADESAO <= TO_DATE(@DATA) + 1 AND OBRIGATORIO = 'S' AND HANDLE <> (SELECT MODULO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = " + SessionVar("HMODBENEFICIARIO") + " ) )"
    End If
  End If

  MODULO.ReadOnly = Not ModuloObrigatorio(SessionVar("HMODBENEFICIARIO"))
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim dllBSBen022 As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String
  Dim SQL As Object
  Set SQL = NewQuery
  Dim SQL2 As Object
  Set SQL2 = NewQuery
  Dim SQL3 As Object
  Set SQL3 = NewQuery
  Set dllBSBen022 = CreateBennerObject("BSBEN022.Modulo")
  Dim viHandle As Long
  Dim viBeneficiario As Long
  Dim viContrato As Long
  Dim vdAdesao As Date

  CanContinue = True

  If ModuloCancelado(SessionVar("HMODBENEFICIARIO")) Then
    bsShowMessage("Módulo já está cancelado!", "E")
    Exit Sub
  End If

  If ModuloObrigatorio(SessionVar("HMODBENEFICIARIO")) Then

    If (CurrentQuery.FieldByName("MODULO").AsInteger = 0) And ModuloObrigatorio(SessionVar("HMODBENEFICIARIO")) Then
      bsShowMessage("O campo módulo é obrigatório! " + CStr(SessionVar("HMODBENEFICIARIO")) + MODULO.Text , "I")
      Exit Sub
    End If

    SQL2.Active = False
    SQL2.Clear
    SQL2.Add("SELECT *                   ")
    SQL2.Add("  FROM SAM_BENEFICIARIO_MOD")
    SQL2.Add(" WHERE HANDLE = :HANDLE    ")
    SQL2.ParamByName("HANDLE").AsInteger = SessionVar("HMODBENEFICIARIO")
    SQL2.Active = True

    viContrato = SQL2.FieldByName("CONTRATO").AsInteger
    viBeneficiario = SQL2.FieldByName("BENEFICIARIO").AsInteger
    vdAdesao = SQL2.FieldByName("DATAADESAO").AsDateTime
    If vdAdesao > CurrentQuery.FieldByName("DATA").AsDateTime Then
		bsShowMessage("Data cancelamento menor que data de adesão", "E")
		CanContinue = False
		Exit Sub
	End If

    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT *                   ")
    SQL.Add("  FROM SAM_BENEFICIARIO_MOD")
    SQL.Add(" WHERE HANDLE = :HANDLE    ")
    SQL.ParamByName("HANDLE").AsInteger = -1
    SQL.RequestLive = True
    SQL.Active = True
    SQL.Insert

    viHandle = NewHandle("SAM_BENEFICIARIO_MOD")
    SQL.FieldByName("HANDLE").AsInteger = viHandle
    SQL.FieldByName("BENEFICIARIO").AsInteger = viBeneficiario
    SQL.FieldByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger

    SQL.FieldByName("CONTRATO").AsInteger = viContrato
    DoAssumir(SQL.TQuery, "SAM_BENEFICIARIO_MOD")
    SQL.FieldByName("DATAADESAO").AsDateTime = CurrentQuery.FieldByName("DATA").AsDateTime + 1

    'validar se já existe o modulo a ser inserido com a mesma data de adesão
	SQL3.Active = False
	SQL3.Clear
	SQL3.Add("SELECT HANDLE FROM SAM_BENEFICIARIO_MOD                                  ")
	SQL3.Add(" WHERE BENEFICIARIO = :BEN And DATAADESAO = :ADESAO And MODULO = :MOD    ")
	SQL3.ParamByName("BEN").AsInteger = viBeneficiario
	SQL3.ParamByName("ADESAO").AsDateTime = SQL.FieldByName("DATAADESAO").AsDateTime
	SQL3.ParamByName("Mod").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
	SQL3.Active = True

	If (SQL3.FieldByName("HANDLE").AsInteger > 0) Then
		bsShowMessage("Já existe este modulo cadastrado para este beneficiario com esta adesão", "E")
		CanContinue = False
		Exit Sub
	End If

    viResultado = dllBSBen022.BeforePost(CurrentSystem, _
                                         SQL2.TQuery, _
                                         True, _
                                         SQL.FieldByName("DATACANCELAMENTO").AsDateTime, _
                                         SQL.FieldByName("MOTIVOCANCELAMENTO").AsInteger, _
                                         vsMensagem)
    If viResultado = 1 Then
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
      SQL.Cancel
    Else
      If vsMensagem <> "" Then
        bsShowMessage(vsMensagem, "I")
      End If
      SQL.Post
    End If

    If CanContinue Then

      SQL.Active = False
      SQL.ParamByName("HANDLE").AsInteger = viHandle
      SQL.Active = True

      viResultado = dllBSBen022.AfterPost(CurrentSystem, _
                                          SQL.TQuery, _
                                          "I", _
                                          "P", _
                                          True, _
                                          vsMensagem)
      If viResultado = 1 Then
        Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
      Else
        If vsMensagem <> "" Then
          bsShowMessage(vsMensagem, "I")
        End If
      End If
    End If

  End If

  viResultado = dllBSBen022.Cancelar(CurrentSystem, _
                                     CLng(SessionVar("HMODBENEFICIARIO")), _
                                     CurrentQuery.FieldByName("DATA").AsDateTime, _
                                     CurrentQuery.FieldByName("MOTIVO").AsInteger, _
                                     vsMensagem)

  If viResultado = 1 Then
    bsShowMessage(vsMensagem, "E")
    CanContinue = False
  Else
    bsShowMessage("Cancelamento concluído!", "I")
  End If

  SQL.Active = False
  Set SQL = Nothing
  SQL2.Active = False
  Set SQL2 = Nothing
  SQL3.Active = False
  Set SQL3 = Nothing
  Set dllBSBen022 = Nothing

  'Se estiver em modo desktop a transação deve ser iniciada
  'Isto é necessário devido ao fato dos componentes BVirtual não controlarem transação
  If CurrentQuery.IsVirtual And _
     VisibleMode And _
     (Not InTransaction) Then
    StartTransaction
  End If
End Sub

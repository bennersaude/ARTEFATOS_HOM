'HASH: BC6D74055330FBB53B98515107D0EB66
'#Uses "*bsShowMessage"

Public Sub BOTAOCALCULAR_OnClick()
  '*******************************************
  'Rotina de referenciamento Calcular'
  '*******************************************
  Dim vRotina
  Dim interface As Object
  Dim Q As Object

  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT GRUPOREFERENCIAMENTO,DATAGERACAO,DATAREFERENCIAMENTO,DATACALCULO FROM SAM_ROTINAREFERENCIAMENTO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.Active = True
  If Not Q.FieldByName("DATAGERACAO").IsNull Then
    If Not Q.FieldByName("DATACALCULO").IsNull Then
      bsShowMessage("Já foi calculado !", "I")
      Set Q = Nothing
      Exit Sub
    Else
      Set interface = CreateBennerObject("BSPRE002.ROTINAS")
      vRotina = RecordHandleOfTable("SAM_ROTINAREFERENCIAMENTO")
      interface.GERAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      RefreshNodesWithTable("SAM_ROTINAREFERENCIAMENTO")
      Set interface = Nothing
      Set Q = Nothing
    End If
  Else
    bsShowMessage("Os candidatos a processamento não foram gerados !", "I")
    Set Q = Nothing
    Exit Sub
  End If

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  '*******************************************
  'Rotina de referenciamento Cancelar'
  '*******************************************
  Dim vRotina
  Dim interface As Object
  Dim Q As Object

  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT GRUPOREFERENCIAMENTO,DATAGERACAO,DATAREFERENCIAMENTO,DATACALCULO FROM SAM_ROTINAREFERENCIAMENTO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.Active = True
  If(Not CurrentQuery.FieldByName("DATAREFERENCIAMENTO").IsNull)Or(Not CurrentQuery.FieldByName("DATACALCULO").IsNull)Or(Not CurrentQuery.FieldByName("DATAGERACAO").IsNull)Then
  Set interface = CreateBennerObject("BSPRE002.ROTINAS")
  vRotina = RecordHandleOfTable("SAM_ROTINAREFERENCIAMENTO")
  interface.CANCELAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  RecordHandleOfTable("SAM_ROTINAREFERENCIAMENTO")
  RefreshNodesWithTable("SAM_ROTINAREFERENCIAMENTO")
  Set interface = Nothing
  Set Q = Nothing
Else
  bsShowMessage("A rotina não foi processada !","I")
  Set Q = Nothing
  Exit Sub
End If
End Sub

Public Sub BOTAOGERAR_OnClick()
  '***************************************
  'Rotina de referenciamento Gerar'
  '***************************************
  Dim vRotina
  Dim interface As Object
  Dim Q As Object
  Dim vHandle As Long


  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT GRUPOREFERENCIAMENTO,DATAGERACAO,DATAREFERENCIAMENTO,DATACALCULO FROM SAM_ROTINAREFERENCIAMENTO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.Active = True

  If Q.FieldByName("DATAREFERENCIAMENTO").IsNull Then
    If CurrentQuery.FieldByName("DATAGERACAO").IsNull Then
      Set interface = CreateBennerObject("BSPRE004.CONSULTAR")
      vHandle = Q.FieldByName("GRUPOREFERENCIAMENTO").AsInteger
      interface.filtro(CurrentSystem, vHandle, "", "T", CurrentQuery.FieldByName("HANDLE").AsInteger)

      REGIMEATENDIMENTO.ReadOnly = True
      LOCALATENDIMENTO.ReadOnly = True
      CONDICAOATENDIMENTO.ReadOnly = True
      TIPOTRATAMENTO.ReadOnly = True
      OBJETIVOTRATAMENTO.ReadOnly = True
      FINALIDADEATENDIMENTO.ReadOnly = True
      CONVENIO.ReadOnly = True
      RefreshNodesWithTable("SAM_ROTINAREFERENCIAMENTO")
    Else
      bsShowMessage("Os candidatos a referenciamento já foram gerados !", "I")
    End If
  Else
    bsShowMessage("Esta rotina já foi processada não pode ser gerada novamente !", "I")
    Set Q = Nothing
    Exit Sub
  End If

End Sub


Public Sub BOTAOREFERENCIAR_OnClick()

  '*******************************************
  'Rotina de referenciamento Referenciar'
  '*******************************************
  Dim vRotina
  Dim interface As Object
  Dim Q As Object

  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT GRUPOREFERENCIAMENTO,DATAGERACAO,DATAREFERENCIAMENTO,DATACALCULO FROM SAM_ROTINAREFERENCIAMENTO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.Active = True
  If Not Q.FieldByName("DATAGERACAO").IsNull Then
    If Not Q.FieldByName("DATACALCULO").IsNull Then
      If Not Q.FieldByName("DATAREFERENCIAMENTO").IsNull Then
        bsShowMessage("Já foi referenciada a rotina !", "I")
        Set Q = Nothing
        Exit Sub
      Else
        Set interface = CreateBennerObject("BSPRE002.ROTINAS")
        vRotina = RecordHandleOfTable("SAM_ROTINAREFERENCIAMENTO")
        interface.PROCESSAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
        RefreshNodesWithTable("SAM_ROTINAREFERENCIAMENTO")
        Set interface = Nothing
        Set Q = Nothing
      End If
    Else
      bsShowMessage("A rotina não foi calculada portanto, ainda não pode ser referenciada !", "I")
      Set Q = Nothing
      Exit Sub
    End If
  Else
    bsShowMessage("Os candidatos a processamento não foram gerados !", "I")
    Set Q = Nothing
    Exit Sub
  End If
End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAGERACAO").IsNull Then
    REGIMEATENDIMENTO.ReadOnly = False
    LOCALATENDIMENTO.ReadOnly = False
    CONDICAOATENDIMENTO.ReadOnly = False
    TIPOTRATAMENTO.ReadOnly = False
    OBJETIVOTRATAMENTO.ReadOnly = False
    FINALIDADEATENDIMENTO.ReadOnly = False
    CONVENIO.ReadOnly = False
  Else
    REGIMEATENDIMENTO.ReadOnly = True
    LOCALATENDIMENTO.ReadOnly = True
    CONDICAOATENDIMENTO.ReadOnly = True
    TIPOTRATAMENTO.ReadOnly = True
    OBJETIVOTRATAMENTO.ReadOnly = True
    FINALIDADEATENDIMENTO.ReadOnly = True
    CONVENIO.ReadOnly = True
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCALCULAR"
			BOTAOCALCULAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOREFERENCIAR"
			BOTAOREFERENCIAR_OnClick
	End Select
End Sub

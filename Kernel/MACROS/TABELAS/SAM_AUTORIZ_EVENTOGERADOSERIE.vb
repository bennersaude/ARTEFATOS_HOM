'HASH: 6141D9AA2846DEAB43790B6FF7C54FA0

'macro SAM_AUTORIZ_EVENTOGERADOSERIE


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("QTDPAGA").AsFloat > 0 Then
    DATADIARIO.ReadOnly = True
    DATAMENSAL.ReadOnly = True
    DATASEMANAL.ReadOnly = True
    EVENTOGERADO.ReadOnly = True
    QTDLIBERADA.ReadOnly = True
    SEMANA.ReadOnly = True
  Else
    DATADIARIO.ReadOnly = False
    DATAMENSAL.ReadOnly = False
    DATASEMANAL.ReadOnly = False
    EVENTOGERADO.ReadOnly = False
    QTDLIBERADA.ReadOnly = False
    SEMANA.ReadOnly = False
  End If

  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT HANDLE FROM SAM_AUTORIZ_EVENTOGERADOSERIE WHERE EVENTOGERADO = :EVENTO")
  Q.ParamByName("EVENTO").Value = RecordHandleOfTable("SAM_AUTORIZ_EVENTOGERADO")
  Q.Active = True

  If Q.EOF Then
    DATADIARIO.ReadOnly = True
    DATAMENSAL.ReadOnly = True
    DATASEMANAL.ReadOnly = True
    EVENTOGERADO.ReadOnly = True
    QTDLIBERADA.ReadOnly = True
    SEMANA.ReadOnly = True
  Else
    DATADIARIO.ReadOnly = False
    DATAMENSAL.ReadOnly = False
    DATASEMANAL.ReadOnly = False
    EVENTOGERADO.ReadOnly = False
    QTDLIBERADA.ReadOnly = False
    SEMANA.ReadOnly = False
  End If

  Set Q = Nothing

  ' --- SMS - 87995 - Início ---------
  Dim SamUtil As Object

  Set SamUtil = CreateBennerObject("SAMUTIL.ROTINAS")
  PERIODODASEMANA.Text = "Período de " + SamUtil.SemanaData(CurrentSystem,Year(CurrentQuery.FieldByName("DATASEMANAL").AsDateTime),CurrentQuery.FieldByName("SEMANA").AsInteger, " até ")
  Set SamUtil = Nothing
  ' --- SMS - 87995 - Fim    ---------

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Q As Object
  Set Q = NewQuery

  If (CurrentQuery.FieldByName("DATADIARIO").IsNull And _
       CurrentQuery.FieldByName("DATASEMANAL").IsNull And _
       CurrentQuery.FieldByName("DATAMENSAL").IsNull) Then
    MsgBox "Data obrigatória."
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("QTDLIBERADA").IsNull Then
    MsgBox "Quantidade liberada obrigatória."
    CanContinue = False
    Exit Sub
  End If

  If ((Not CurrentQuery.FieldByName("DATASEMANAL").IsNull) And _
      (CurrentQuery.FieldByName("SEMANA").IsNull)) Then
    MsgBox "Semana obrigatória."
    CanContinue = False
    Exit Sub
  End If

  Select Case TIPOPERIODO.PageIndex + 1
    Case 1
      Q.Clear
      Q.Add("SELECT 'S' EXISTE")
      Q.Add("  FROM SAM_AUTORIZ_EVENTOGERADOSERIE")
      Q.Add(" WHERE DATADIARIO = :DATA AND HANDLE <> :HANDLE AND EVENTOGERADO = :EVENTO")
      Q.ParamByName("DATA").Value = CurrentQuery.FieldByName("DATADIARIO").AsDateTime
      Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOGERADO").AsInteger
      Q.Active = True
      CurrentQuery.FieldByName("DATASEMANAL").Clear
      CurrentQuery.FieldByName("DATAMENSAL").Clear
    Case 2
      Q.Clear
      Q.Add("SELECT 'S' EXISTE")
      Q.Add("  FROM SAM_AUTORIZ_EVENTOGERADOSERIE")
      Q.Add(" WHERE DATASEMANAL = :DATA AND SEMANA = :SEMANA AND HANDLE <> :HANDLE AND EVENTOGERADO = :EVENTO")
      Q.ParamByName("DATA").Value = CurrentQuery.FieldByName("DATASEMANAL").AsDateTime
      Q.ParamByName("SEMANA").Value = CurrentQuery.FieldByName("SEMANA").AsInteger
      Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOGERADO").AsInteger
      CurrentQuery.FieldByName("DATADIARIO").Clear
      CurrentQuery.FieldByName("DATAMENSAL").Clear
      Q.Active = True
    Case 3
      Q.Clear
      Q.Add("SELECT 'S' EXISTE")
      Q.Add("  FROM SAM_AUTORIZ_EVENTOGERADOSERIE")
      Q.Add(" WHERE DATAMENSAL = :DATA AND HANDLE <> :HANDLE AND EVENTOGERADO = :EVENTO")
      Q.ParamByName("DATA").Value = CurrentQuery.FieldByName("DATAMENSAL").AsDateTime
      Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOGERADO").AsInteger
      Q.Active = True
      CurrentQuery.FieldByName("DATADIARIO").Clear
      CurrentQuery.FieldByName("DATASEMANAL").Clear
  End Select

  If Not Q.EOF Then
    MsgBox "Existe uma seção nessa data."
    CanContinue = False
    Q.Active = False
    Set Q = Nothing
    Exit Sub
  End If

  Q.Clear
  Q.Add("SELECT TIPOPERIODO")
  Q.Add("  FROM SAM_AUTORIZ_EVENTOGERADOSERIE")
  Q.Add(" WHERE HANDLE <> :HANDLE AND EVENTOGERADO = :EVENTO AND TIPOPERIODO <> :TIPOPERIODO")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOGERADO").AsInteger
  Q.ParamByName("TIPOPERIODO").Value = CurrentQuery.FieldByName("TIPOPERIODO").AsInteger
  Q.Active = True

  If Not Q.EOF Then
    MsgBox "Um evento não pode ter tipo de períodos diferentes."
    CanContinue = False
    Q.Active = False
    Set Q = Nothing
    Exit Sub
  End If

  Set Q = Nothing

End Sub


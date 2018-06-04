'HASH: 81905F2D20763330B6155F154072FE07
'#Uses "*bsShowMessage"
'CLI_ROTALTERARECURSO

Public Sub MontaRotulo()
  ROTULODIASEMANA.Text = ""

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT DIASEMANA, TIPOATIVIDADE FROM CLI_ESCALA WHERE HANDLE = :ESCALA")
  SQL.ParamByName("ESCALA").Value = CurrentQuery.FieldByName("ESCALASUBSTITUIDO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("DIASEMANA").AsInteger = 1 Then
    ROTULODIASEMANA.Text = "Dia da semana: Domingo"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 2 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 3 Then
    ROTULODIASEMANA.Text = "Dia da semana: Terça"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 4 Then
    ROTULODIASEMANA.Text = "Dia da semana: Quarta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 5 Then
    ROTULODIASEMANA.Text = "Dia da semana: Quinta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 6 Then
    ROTULODIASEMANA.Text = "Dia da semana: Sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 7 Then
    ROTULODIASEMANA.Text = "Dia da semana: Sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 8 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 9 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 10 Then
    ROTULODIASEMANA.Text = "Dia da semana: Segunda a Domingo"
  End If

  If SQL.FieldByName("TIPOATIVIDADE").AsString = "A" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Ambos"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "P" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Procedimento"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "C" Then
    ROTULODIASEMANA.Text = ROTULODIASEMANA.Text + "   Atividade: Consulta"
  End If

  SQL.ParamByName("ESCALA").Value = CurrentQuery.FieldByName("ESCALASUBSTITUTO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("DIASEMANA").AsInteger = 1 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Domingo"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 2 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Segunda"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 3 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Terça"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 4 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Quarta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 5 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Quinta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 6 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 7 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 8 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Segunda a sexta"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 9 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Segunda a sábado"
  End If
  If SQL.FieldByName("DIASEMANA").AsInteger = 10 Then
    ROTULODIASEMANA2.Text = "Dia da semana: Segunda a Domingo"
  End If

  If SQL.FieldByName("TIPOATIVIDADE").AsString = "A" Then
    ROTULODIASEMANA2.Text = ROTULODIASEMANA2.Text + "   Atividade: Ambos"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "P" Then
    ROTULODIASEMANA2.Text = ROTULODIASEMANA2.Text + "   Atividade: Procedimento"
  End If
  If SQL.FieldByName("TIPOATIVIDADE").AsString = "C" Then
    ROTULODIASEMANA2.Text = ROTULODIASEMANA2.Text + "   Atividade: Consulta"
  End If

  Set SQL = Nothing
End Sub

Public Sub BOTAOALTERAR_OnClick()
  If bsShowMessage("Confirma a alteração dos recursos?", "Q") = vbYes Then



  	Dim SQL As Object
  	Set SQL = NewQuery
  	If CurrentQuery.FieldByName("USUARIOALTEROU").IsNull Then
	    SQL.Add("SELECT HANDLE FROM CLI_ROTALTERARECURSOAGENDA")
	    SQL.Add(" WHERE ROTALTERARECURSO = :ROTINA")
	    SQL.Add("   AND SITUACAO <> 'P'")
	    SQL.Add("UNION")
	    SQL.Add("SELECT HANDLE FROM CLI_ROTALTERARECURSOAGENDACONF")
	    SQL.Add(" WHERE ROTALTERARECURSO = :ROTINA")
	    SQL.Add("   AND SITUACAO <> 'P'")
	    SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	    SQL.Active = True

    	If Not SQL.EOF Then
      	Dim BSCLI001DLL As Object
      	Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
      	BSCLI001DLL.ALTERARECURSO(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "A")
      	Set BSCLI001DLL = Nothing
      	RefreshNodesWithTable("CLI_ROTALTERARECURSO")
    	Else
      	bsShowMessage("Não existe nenhuma consulta a ser alterada!", "I")
      	Exit Sub
	    End If
  	Else
	    bsShowMessage("Alteração já foi executada!", "I")
	    Exit Sub
  	End If
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.FieldByName("USUARIOALTEROU").IsNull Then
    Dim BSCLI001DLL As Object
    Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
    BSCLI001DLL.ALTERARECURSO(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")
    Set BSCLI001DLL = Nothing

    RefreshNodesWithTable("CLI_ROTALTERARECURSO")
  Else
    bsShowMessage("Alteração já foi executada, não é mais permitido cancelar!", "I")
    Exit Sub
  End If
End Sub

Public Sub BOTAOCONFIRMARTUDO_OnClick()
  If CurrentQuery.FieldByName("USUARIOALTEROU").IsNull Then
    Dim SQL As Object
    Set SQL = NewQuery
    If Not InTransaction Then StartTransaction

    SQL.Clear
    SQL.Add("UPDATE CLI_ROTALTERARECURSOAGENDA SET SITUACAO = 'C'")
    SQL.Add("WHERE ROTALTERARECURSO = :ROTINA AND SITUACAO = 'P'")
    SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    SQL.Clear
    SQL.Add("UPDATE CLI_ROTALTERARECURSOAGENDACONF SET SITUACAO = 'C'")
    SQL.Add("WHERE ROTALTERARECURSO = :ROTINA AND SITUACAO = 'P'")
    SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    If InTransaction Then Commit
    Set SQL = Nothing
    RefreshNodesWithTable("CLI_ROTALTERARECURSO")
  Else
    bsShowMessage("Alteração já foi executada, não é mais permitido confirmar tudo!", "I")
    Exit Sub
  End If
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT 1 FROM CLI_ROTALTERARECURSOAGENDA WHERE ROTALTERARECURSO = :ROTINA")
  SQL.Add("UNION")
  SQL.Add("SELECT 1 FROM CLI_ROTALTERARECURSOAGENDACONF WHERE ROTALTERARECURSO = :ROTINA")
  SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Rotina já foi gerada!", "E")
    CanContinue = False
    Exit Sub
  Else
    Dim BSCLI001DLL As Object
    Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
    BSCLI001DLL.ALTERARECURSO(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "G")
    Set BSCLI001DLL = Nothing
  End If
  Set SQL = Nothing
End Sub

Public Sub ESCALASUBSTITUIDO_OnChange()
  MontaRotulo
End Sub


Public Sub ESCALASUBSTITUTO_OnChange()
  MontaRotulo
End Sub

Public Sub TABLE_AfterScroll()
  MontaRotulo
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If WebMode Then
  	 ESCALASUBSTITUTO.WebLocalWhere = "A.DISPONIVEL = 'S'"
  	 ESCALASUBSTITUIDO.WebLocalWhere = "A.DISPONIVEL = 'S'"
  ElseIf VisibleMode Then
  	 ESCALASUBSTITUTO.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
     ESCALASUBSTITUIDO.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
  End If



  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT 1 FROM CLI_ROTALTERARECURSOAGENDA WHERE ROTALTERARECURSO = :ROTINA")
  SQL.Add("UNION")
  SQL.Add("SELECT 1 FROM CLI_ROTALTERARECURSOAGENDACONF WHERE ROTALTERARECURSO = :ROTINA")
  SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    bsShowMessage("Rotina não pode ser alterada!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	 ESCALASUBSTITUTO.WebLocalWhere = "A.DISPONIVEL = 'S'"
  	 ESCALASUBSTITUIDO.WebLocalWhere = "A.DISPONIVEL = 'S'"
  ElseIf VisibleMode Then
  	 ESCALASUBSTITUTO.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
     ESCALASUBSTITUIDO.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
    bsShowMessage("A data final não pode ser anterior a data inicial!", "I")
    CanContinue = False
    Exit Sub
  End If

  Dim ClinicaDLL As Object
  Set ClinicaDLL = CreateBennerObject("CliClinica.Agenda")
  If Not ClinicaDLL.EscalaMaior(CurrentSystem, CurrentQuery.FieldByName("ESCALASUBSTITUIDO").AsInteger, _
                                 CurrentQuery.FieldByName("ESCALASUBSTITUTO").AsInteger, _
                                 CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                                 CurrentQuery.FieldByName("DATAFINAL").AsDateTime)Then
    bsShowMessage("Escalas Incompatíveis!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set ClinicaDLL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOALTERAR") Then
		BOTAOALTERAR_OnClick
	End If
	If (CommandID = "BOTAOCANCELAR") Then
		BOTAOCANCELAR_OnClick
	End If
	If (CommandID = "BOTAOCONFIRMARTUDO") Then
		BOTAOCONFIRMARTUDO_OnClick
	End If
	If (CommandID = "BOTAOGERAR") Then
		BOTAOGERAR_OnClick
	End If
End Sub

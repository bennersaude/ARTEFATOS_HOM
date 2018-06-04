'HASH: AB0F37729D17FB9FE64935A7CE47EAFE
'Macro: SAM_LIMITACAO
'#Uses "*bsShowMessage"

Dim viPeriodicidadeAnterior As Long

Public Sub TABLE_AfterScroll()
  BOTAOGERAREVENTOS.Visible =False
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	viPeriodicidadeAnterior = CurrentQuery.FieldByName("PERIODICIDADE").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (CurrentQuery.State = 2) And (viPeriodicidadeAnterior <> CurrentQuery.FieldByName("PERIODICIDADE").AsInteger) Then




 	Dim SQL As Object
  	Set SQL = NewQuery

  	SQL.Active = False
  	SQL.Clear
  	SQL.Add("SELECT COUNT(1) QTDE FROM SAM_CONTRATO_CONTLIM WHERE LIMITACAO = :HANDLE")
  	SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  	SQL.Active = True

  	If (SQL.FieldByName("QTDE").AsInteger = 0) Then
	    SQL.Active = False
	    SQL.Clear
	    SQL.Add("SELECT COUNT(1) QTDE FROM SAM_FAMILIA_CONTLIM WHERE LIMITACAO = :HANDLE")
	    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	    SQL.Active = True

    	If (SQL.FieldByName("QTDE").AsInteger = 0) Then
      	SQL.Active = False
      	SQL.Clear
      	SQL.Add("SELECT COUNT(1) QTDE FROM SAM_BENEFICIARIO_CONTLIM WHERE LIMITACAO = :HANDLE")
      	SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      	SQL.Active = True

      	If (SQL.FieldByName("QTDE").AsInteger > 0) Then
        	bsShowMessage("Não é permitido alterar a periodicidade de limitação com contagem associada!", "E")
	        CanContinue = False
      	End If
    	Else
      	bsShowMessage("Não é permitido alterar a periodicidade de limitação com contagem associada!", "E")
      	CanContinue = False
	    End If
  	Else
	    bsShowMessage("Não é permitido alterar a periodicidade de limitação com contagem associada!", "E")
	    CanContinue = False
  	End If
  	Set SQL = Nothing


  End If
End Sub

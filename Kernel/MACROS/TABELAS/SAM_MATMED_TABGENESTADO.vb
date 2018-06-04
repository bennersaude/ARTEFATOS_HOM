'HASH: 2ED060E0FD9A6089647DD23C2755BCF4
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim SQLEstado As Object
Set SQLEstado = NewQuery

  If CurrentQuery.FieldByName("TABIMPORTAFRACIONADO").AsString = "2" Then
	If CurrentQuery.FieldByName("PRECOFRACIONADOMARCA").IsNull Then
		bsShowMessage("Tabela genérica Fracionada de marca não preenchida!", "E")
	      	CanContinue = False
        Exit Sub
	End If
	If CurrentQuery.FieldByName("PRECOFRACIONADOGENERICO").IsNull Then
		BsShowMessage("Tabela genérica Fracionada de medicamentos genéricos não preenchida!", "E")
	      	CanContinue = False
	End If
	If CurrentQuery.FieldByName("PRECOFRACIONADOMARCARESTHOSP").IsNull Then
		BsShowMessage("Tabela genérica Fracionada de marca não preenchida!", "E")
	      	CanContinue = False
	End If
	If CurrentQuery.FieldByName("PRECOFRACIONADOGENERICORESTHOS").IsNull Then
		BsShowMessage("Tabela genérica Fracionada de marca não preenchida!", "E")
	      	CanContinue = False
	End If
  End If

  If CurrentQuery.FieldByName("ESTADO").AsString <> "ZF" Then

    SQLEstado.Add("SELECT COUNT(1) QTD FROM ESTADOS WHERE SIGLA = :SIGLA")
    SQLEstado.ParamByName("SIGLA").AsString = CurrentQuery.FieldByName("ESTADO").AsString
    SQLEstado.Active = True

    If SQLEstado.FieldByName("QTD").AsInteger = 0 Then
      BsShowMessage("A sigla digitada é inválida!", "E")
      CanContinue = False
    End If
  End If

  SQLEstado.Clear
  SQLEstado.Add("SELECT COUNT(1) QTD FROM SAM_MATMED_TABGENESTADO WHERE ESTADO = :ESTADO AND HANDLE <> :HANDLE")
  SQLEstado.ParamByName("ESTADO").AsString = CurrentQuery.FieldByName("ESTADO").AsString
  SQLEstado.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLEstado.Active = True

  If SQLEstado.FieldByName("QTD").AsInteger > 0 Then
    BsShowMessage("Já existe parâmetros para este estado!", "E")
    CanContinue = False
  End If

  Set SQLEstado = Nothing

End Sub

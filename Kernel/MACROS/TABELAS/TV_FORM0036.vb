'HASH: A9F6DD792113B660575E5472A396991B
 '#Uses "*bsShowMessage"

  Option Explicit

  Dim query As Object
  Dim Query1 As Object
  Dim Query2 As Object
  Dim Query3 As Object
  Dim Query4 As Object
  Dim Query5 As Object
  Dim Query6 As Object
  Dim QueryDestino As Object

  Dim HandleInserido_Aux As Long
  Dim HandleInserido_Aux1 As Long
  Dim HandleInserido_Aux2 As Long
  Dim HandleInserido_Aux3 As Long
  Dim HandleInserido_Aux4 As Long
  Dim HandleInserido_Aux5 As Long
  Dim HandleInserido_Aux6 As Long

  Dim handleinserido As Long
  Dim handleinserido1 As Long
  Dim handleinserido2 As Long
  Dim handleinserido3 As Long
  Dim handleinserido4 As Long
  Dim handleinserido5 As Long
  Dim handleinserido6 As Long
  Dim HandleInseridoContab As Long
  Dim ValorEmpresaMPU As String


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  ValorEmpresaMPU = ""
  If SessionVar("VALOREMPRESAMPU") <> "" Then
    ValorEmpresaMPU = SessionVar("VALOREMPRESAMPU")
  End If

  Set query = NewQuery
  Set Query1 = NewQuery
  Set Query2 = NewQuery
  Set Query3 = NewQuery
  Set Query4 = NewQuery
  Set Query5 = NewQuery
  Set Query6 = NewQuery
  Set QueryDestino = NewQuery

  query.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE ESTRUTURA = :PESTRUTURA")
  query.ParamByName("PESTRUTURA").AsString = CurrentQuery.FieldByName("ESTRUTURAORIGEM").AsString
  query.Active = True

  If query.EOF Then
    bsShowMessage("A Classe Gerencial informada não foi encontrada !", "I")
    Exit Sub
  End If

  QueryDestino.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE ESTRUTURA = :PESTRUTURA")
  QueryDestino.ParamByName("PESTRUTURA").AsString = CurrentQuery.FieldByName("ESTRUTURADESTINO").AsString
  QueryDestino.Active = True

  If QueryDestino.EOF Then
    bsShowMessage("A estrutura de Destino não foi encontrada!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ESTRUTURADESTINO").AsString = CurrentQuery.FieldByName("ESTRUTURAORIGEM").AsString Then
    bsShowMessage("A classe de origem e de destino são as mesmas !", "I")
    Exit Sub
  End If

  HandleInserido_Aux = QueryDestino.FieldByName("HANDLE").AsInteger

  Query1.Clear
  Query1.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
  Query1.ParamByName("PHANDLE").AsInteger = query.FieldByName("HANDLE").AsInteger
  Query1.Active = True

  While Not Query1.EOF
    Insere1 Query1, "N"
    If Query1.FieldByName("ULTIMONIVEL").AsString = "S" Then
      InsereContabilizacao(Query1)
    Else
      Query2.Clear
      Query2.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
      Query2.ParamByName("PHANDLE").AsInteger = Query1.FieldByName("HANDLE").AsInteger
      Query2.Active = True

      While Not Query2.EOF
        Insere2 Query2, "N"
        If Query2.FieldByName("ULTIMONIVEL").AsString = "S" Then
          InsereContabilizacao(Query2)
        Else
          Query3.Clear
          Query3.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
          Query3.ParamByName("PHANDLE").AsInteger = Query2.FieldByName("HANDLE").AsInteger
          Query3.Active = True

          While Not Query3.EOF
            Insere3 Query3, "N"
            If Query3.FieldByName("ULTIMONIVEL").AsString = "S" Then
              InsereContabilizacao(Query3)
            Else
              Query4.Clear
              Query4.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
              Query4.ParamByName("PHANDLE").AsInteger = Query3.FieldByName("HANDLE").AsInteger
              Query4.Active = True

              While Not Query4.EOF
                Insere4 Query4, "N"
                If Query4.FieldByName("ULTIMONIVEL").AsString = "S" Then
                  InsereContabilizacao(Query4)
                Else
                  Query5.Clear
                  Query5.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
                  Query5.ParamByName("PHANDLE").AsInteger = Query4.FieldByName("HANDLE").AsInteger
                  Query5.Active = True

                  While Not Query5.EOF
                    Insere5 Query5, "N"
                    If Query5.FieldByName("ULTIMONIVEL").AsString = "S" Then
                      InsereContabilizacao(Query5)
                    Else
                      Query6.Clear
                      Query6.Add("SELECT * FROM SFN_CLASSEGERENCIAL WHERE NIVELSUPERIOR = :PHANDLE")
                      Query6.ParamByName("PHANDLE").AsInteger = Query5.FieldByName("HANDLE").AsInteger
                      Query6.Active = True

                      While Not Query6.EOF
                        Insere6 Query6, "N"
                        If Query6.FieldByName("ULTIMONIVEL").AsString = "S" Then
                          InsereContabilizacao(Query6)
                        End If
                        Query6.Next
                      Wend

                    End If
                    Query5.Next
                  Wend

                End If
                Query4.Next
              Wend

            End If
            Query3.Next
          Wend
        End If
        Query2.Next
      Wend
    End If
    Query1.Next
  Wend
End Sub

Public Sub Insere(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  If Not InTransaction Then StartTransaction
  	qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
	qInsert.Add(")")
  	qInsert.Add("VALUES")
  	qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If

  	qInsert.Add(")")

  	handleinserido = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido
  	'qInsert.ParamByName("ESTRUTURA").AsString = Replace(query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString, vEstrutura)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(query.FieldByName("ESTRUTURA").AsString, CurrentQuery.FieldByName("ESTRUTURADESTINO").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = QuerySelecao.FieldByName("NIVELSUPERIOR").AsInteger
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

  	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux = handleinserido
  HandleInseridoContab = handleinserido

  Set qInsert = Nothing

End Sub

Public Sub Insere1(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
  If ValorEmpresaMPU <> "" Then
    qInsert.Add(", EMPRESA ")
  End If
  qInsert.Add(")")
  qInsert.Add("VALUES")
  qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
  If ValorEmpresaMPU <> "" Then
    qInsert.Add(", :EMPRESA ")
  End If

  qInsert.Add(")")


  handleinserido1 = NewHandle("SFN_CLASSEGERENCIAL")

  qInsert.ParamByName("HANDLE").AsInteger = handleinserido1
  'qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  'Coelho SMS: 68853 - incluida a function SubstEstrutura
  qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido1
  qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
  qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux
  If ValorEmpresaMPU <> "" Then
    qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
  End If

  qInsert.ExecSQL
  HandleInserido_Aux1 = handleinserido1
  HandleInseridoContab = handleinserido1

  Set qInsert = Nothing

End Sub

Public Sub Insere2(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  If Not InTransaction Then StartTransaction
  	qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
    qInsert.Add(")")

  	qInsert.Add("VALUES")
  	qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If
    qInsert.Add(")")

  	handleinserido2 = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido2
  	' qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido2
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux1
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

  	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux2 = handleinserido2
  HandleInseridoContab = handleinserido2

  Set qInsert = Nothing

End Sub

Public Sub Insere3(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  If Not InTransaction Then StartTransaction
    qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
    qInsert.Add(")")

  	qInsert.Add("VALUES")
  	qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If
    qInsert.Add(")")


  	handleinserido3 = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido3
  	' qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido3
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux2
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

  	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux3 = handleinserido3
  HandleInseridoContab = handleinserido3

  Set qInsert = Nothing

End Sub

Public Sub Insere4(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  If Not InTransaction Then StartTransaction
  	qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
    qInsert.Add(")")

  	qInsert.Add("VALUES")
  	qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If
    qInsert.Add(")")


  	handleinserido4 = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido4
  	'qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido4
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux3
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux4 = handleinserido4
  HandleInseridoContab = handleinserido4

  Set qInsert = Nothing

End Sub

Public Sub Insere5(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery
  If Not InTransaction Then StartTransaction
  	qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
    qInsert.Add(")")

  	qInsert.Add("VALUES")
    qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If
    qInsert.Add(")")


  	handleinserido5 = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido5
  	' qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido5
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux4
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

  	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux5 = handleinserido5
  HandleInseridoContab = handleinserido5

  Set qInsert = Nothing

End Sub

Public Sub Insere6(QuerySelecao As Object, Principal As String)
  Dim qInsert As Object
  Set qInsert = NewQuery

  If Not InTransaction Then StartTransaction
	qInsert.Add("INSERT INTO SFN_CLASSEGERENCIAL")
  	qInsert.Add(" (HANDLE,  ESTRUTURA, CODIGOREDUZIDO, ULTIMONIVEL, DESCRICAO, NATUREZA, NIVELSUPERIOR, TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", EMPRESA ")
    End If
    qInsert.Add(")")

  	qInsert.Add("VALUES")
  	qInsert.Add(" (:HANDLE,  :ESTRUTURA, :CODIGOREDUZIDO, :ULTIMONIVEL, :DESCRICAO, :NATUREZA, :NIVELSUPERIOR, :TIPOATO ")
    If ValorEmpresaMPU <> "" Then
      qInsert.Add(", :EMPRESA ")
    End If
    qInsert.Add(")")


  	handleinserido6 = NewHandle("SFN_CLASSEGERENCIAL")

  	qInsert.ParamByName("HANDLE").AsInteger = handleinserido6
  	'qInsert.ParamByName("ESTRUTURA").AsString = Replace(QuerySelecao.FieldByName("ESTRUTURA").AsString, query.FieldByName("ESTRUTURA").AsString, QueryDestino.FieldByName("ESTRUTURA").AsString)
  	'Coelho SMS: 68853 - incluida a function SubstEstrutura
  	qInsert.ParamByName("ESTRUTURA").AsString = SubstEstrutura(QuerySelecao.FieldByName("ESTRUTURA").AsString,QueryDestino.FieldByName("ESTRUTURA").AsString)
  	qInsert.ParamByName("CODIGOREDUZIDO").AsInteger = handleinserido6
  	qInsert.ParamByName("ULTIMONIVEL").AsString = QuerySelecao.FieldByName("ULTIMONIVEL").AsString
  	qInsert.ParamByName("DESCRICAO").AsString = QuerySelecao.FieldByName("DESCRICAO").AsString
  	qInsert.ParamByName("NATUREZA").AsString = QuerySelecao.FieldByName("NATUREZA").AsString
  	qInsert.ParamByName("TIPOATO").AsString = QuerySelecao.FieldByName("TIPOATO").AsString
  	qInsert.ParamByName("NIVELSUPERIOR").AsInteger = HandleInserido_Aux5
    If ValorEmpresaMPU <> "" Then
      qInsert.ParamByName("EMPRESA").AsInteger = CInt(ValorEmpresaMPU)
    End If

  	qInsert.ExecSQL
  If InTransaction Then Commit
  HandleInserido_Aux6 = handleinserido6
  HandleInseridoContab = handleinserido6

  Set qInsert = Nothing

End Sub

Public Sub InsereContabilizacao(QuerySelecao As Object)
  Dim qInsert As Object
  Set qInsert = NewQuery

  Dim QueryContabilizacao As Object
  Set QueryContabilizacao = NewQuery

  QueryContabilizacao.Add("SELECT * FROM SFN_CONTABILIZACAO WHERE CLASSEGERENCIAL =:PHANDLE")
  QueryContabilizacao.ParamByName("PHANDLE").AsInteger = QuerySelecao.FieldByName("HANDLE").AsInteger
  QueryContabilizacao.Active = True

  If Not InTransaction Then StartTransaction

  qInsert.Add("INSERT INTO SFN_CONTABILIZACAO")
  qInsert.Add("(CLASSECONTABILCRE,")
  qInsert.Add(" CLASSECONTABILDEB,")
  qInsert.Add(" CLASSEGERENCIAL,")
  qInsert.Add(" CONTABHIST,")
  qInsert.Add(" HANDLE,")
  qInsert.Add(" OPERACAO,")
  qInsert.Add(" TABCLASSECRE,")
  qInsert.Add(" TABCLASSEDEB )")
  qInsert.Add("VALUES")
  qInsert.Add("(:CLASSECONTABILCRE,")
  qInsert.Add(" :CLASSECONTABILDEB,")
  qInsert.Add(" :CLASSEGERENCIAL,")
  qInsert.Add(" :CONTABHIST,")
  qInsert.Add(" :HANDLE,")
  qInsert.Add(" :OPERACAO,")
  qInsert.Add(" :TABCLASSECRE,")
  qInsert.Add(" :TABCLASSEDEB )")

  While Not QueryContabilizacao.EOF

    If QueryContabilizacao.FieldByName("CLASSECONTABILCRE").IsNull Then
      qInsert.ParamByName("CLASSECONTABILCRE").DataType = ftInteger
      qInsert.ParamByName("CLASSECONTABILCRE").Clear
    Else
      qInsert.ParamByName("CLASSECONTABILCRE").AsInteger = QueryContabilizacao.FieldByName("CLASSECONTABILCRE").AsInteger
    End If

    If QueryContabilizacao.FieldByName("CLASSECONTABILDEB").IsNull Then
      qInsert.ParamByName("CLASSECONTABILDEB").DataType = ftInteger
      qInsert.ParamByName("CLASSECONTABILDEB").Clear
    Else
      qInsert.ParamByName("CLASSECONTABILDEB").AsInteger = QueryContabilizacao.FieldByName("CLASSECONTABILDEB").AsInteger
    End If

    qInsert.Active = False
    qInsert.ParamByName("CLASSEGERENCIAL").AsInteger = HandleInseridoContab
    qInsert.ParamByName("CONTABHIST").AsInteger = QueryContabilizacao.FieldByName("CONTABHIST").AsInteger
    qInsert.ParamByName("HANDLE").AsInteger = NewHandle("SFN_CONTABILIZACAO")
    qInsert.ParamByName("OPERACAO").AsInteger = QueryContabilizacao.FieldByName("OPERACAO").AsInteger

    If QueryContabilizacao.FieldByName("TABCLASSECRE").IsNull Then
      qInsert.ParamByName("TABCLASSECRE").DataType = ftInteger
      qInsert.ParamByName("TABCLASSECRE").Clear
    Else
      qInsert.ParamByName("TABCLASSECRE").AsInteger = QueryContabilizacao.FieldByName("TABCLASSECRE").AsInteger
    End If

    If QueryContabilizacao.FieldByName("TABCLASSEDEB").IsNull Then
      qInsert.ParamByName("TABCLASSEDEB").DataType = ftInteger
      qInsert.ParamByName("TABCLASSEDEB").Clear
    Else
      qInsert.ParamByName("TABCLASSEDEB").AsInteger = QueryContabilizacao.FieldByName("TABCLASSEDEB").AsInteger
    End If

    qInsert.ExecSQL

    QueryContabilizacao.Next

  Wend
  If InTransaction Then Commit
  Set qInsert = Nothing
  Set QueryContabilizacao = Nothing
End Sub

' Coelho SMS: 68853
Public Function SubstEstrutura(EstOriginal As String, CharSubst As String) As String
  SubstEstrutura = Mid(EstOriginal, Len(CharSubst) + 1, Len(EstOriginal))
  SubstEstrutura = CharSubst + SubstEstrutura
End Function

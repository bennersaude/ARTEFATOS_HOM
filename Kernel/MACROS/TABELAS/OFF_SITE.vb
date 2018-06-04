'HASH: 8D68F379F60F577ECE4CC80AA9806820
 
'#Uses "*bsShowMessage"
Public Sub PROCESSAR_OnClick()
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Active = False
Consulta.Add("SELECT *                   ")
Consulta.Add("  FROM OFF_SITELOCAL       ")
Consulta.Active = True

If Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Informar o site local!", "E")
  CanContinue = False
  Exit Sub
End If

Consulta.Clear
Consulta.Active = False
Consulta.Add("SELECT *                   ")
Consulta.Add("  FROM OFF_SITELOCAL       ")
Consulta.Add(" WHERE NOMESITE = :CORRENTE")
Consulta.ParamByName("CORRENTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
Consulta.Active = True

If Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Só é possível processar o site local!", "E")
  CanContinue = False
  Exit Sub
End If



Dim Interface As Object

If VisibleMode Then
	Set Interface = CreateBennerObject("BSInterfaceOffLine.Geral")

	Interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
Else
	Set Interface = CreateBennerObject("OFFLINE.Geral")
	Interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	bsShowMessage("Opa", "I")
End If

RefreshNodesWithTable("OFF_SITE")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

If CurrentQuery.FieldByName("NOMESITE").IsNull Then
  bsShowMessage("Nome do site está vazio!", "E")
  CanContinue = False
  Exit Sub
End If

If (CurrentQuery.FieldByName("VALORMINIMO").IsNull) _
   Or (CurrentQuery.FieldByName("VALORMINIMO").Value <=0 ) Then
  bsShowMessage("Valor mínimo está incorreto!", "I")
  CanContinue = False
  Exit Sub
End If

If (CurrentQuery.FieldByName("VALORMINIMO").Value > CurrentQuery.FieldByName("VALORMAXIMO").Value) _
  And (CurrentQuery.FieldByName("VALORMAXIMO").Value > 0)  Then
  bsShowMessage("Valor mínimo está maior que o valor máximo!", "E")
  CanContinue = False
  Exit Sub
  
End If

Dim ConsultaSite As Object
Set ConsultaSite = NewQuery

ConsultaSite.Clear
ConsultaSite.Active = False
ConsultaSite.Add("SELECT COUNT(NOMESITE) QTD ")
ConsultaSite.Add("  FROM OFF_SITE            ")
ConsultaSite.Add("  WHERE SITECENTRAL = 'S'  ")
ConsultaSite.Active = True

If ConsultaSite.FieldByName("QTD").AsInteger > 0 Then
  bsShowMessage("Já existe um outro Site consfigurado como Central", "E")
  CanContinue = False
  Exit Sub
End If

Dim Consulta As Object
Set Consulta = NewQuery

If CurrentQuery.FieldByName("VALORMAXIMO").IsNull Then
  Consulta.Clear
  Consulta.Active = False
  Consulta.Add("SELECT HANDLE                       ")
  Consulta.Add("  FROM OFF_SITE                     ")
  Consulta.Add(" WHERE VALORMAXIMO IS NULL          ")
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    Consulta.Add(" AND HANDLE <> :HANDLE            ")
    Consulta.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  End If
  Consulta.Active = True

  If Not Consulta.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Existe cruzamento de sequence!", "E")
    CanContinue = False
    Exit Sub
  End If


  Consulta.Clear
  Consulta.Active = False
  Consulta.Add("SELECT VALORMINIMO                  ")
  Consulta.Add("  FROM OFF_SITE                     ")
  Consulta.Add(" WHERE VALORMAXIMO >= :QTDEMIN      ")
  Consulta.ParamByName("QTDEMIN").AsInteger = CurrentQuery.FieldByName("VALORMINIMO").AsInteger
  Consulta.Active = True

  If Not Consulta.FieldByName("VALORMINIMO").IsNull Then
    bsShowMessage("Existe cruzamento de sequence!", "E")
    CanContinue = False
    Exit Sub
  End If
Else
  Consulta.Clear
  Consulta.Active = False
  Consulta.Add("SELECT VALORMINIMO, VALORMAXIMO                                  ")
  Consulta.Add("  FROM OFF_SITE                                                  ")
  Consulta.Add(" WHERE (((:QTDEMIN BETWEEN (VALORMINIMO) AND (VALORMAXIMO))  ")
  Consulta.Add("    OR   (:QTDEMAX BETWEEN (VALORMINIMO) AND (VALORMAXIMO))) ")
  Consulta.Add("    OR  (( (VALORMINIMO) BETWEEN :QTDEMIN AND :QTDEMAX)  ")
  Consulta.Add("    OR   ( (VALORMAXIMO) BETWEEN :QTDEMIN AND :QTDEMAX)))  ")
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    Consulta.Add(" AND HANDLE <> :HANDLE")
    Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
  Consulta.ParamByName("QTDEMIN").AsInteger = CurrentQuery.FieldByName("VALORMINIMO").AsInteger
  Consulta.ParamByName("QTDEMAX").AsInteger = CurrentQuery.FieldByName("VALORMAXIMO").AsInteger
  Consulta.Active = True

  If Not Consulta.FieldByName("VALORMINIMO").IsNull Or Not Consulta.FieldByName("VALORMAXIMO").IsNull Then
    bsShowMessage("Existe cruzamento de sequence!", "E")
    CanContinue = False
    Exit Sub
  End If
End If



End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "PROCESSAR" Then
		PROCESSAR_OnClick
	End If
End Sub

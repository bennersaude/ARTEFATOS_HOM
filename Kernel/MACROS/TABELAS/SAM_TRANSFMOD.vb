'HASH: 51604BDF754424267DEA599C2B1D0CFF
 
 Public Sub MODULODESTINO_OnPopup(ShowPopup As Boolean)
  ShowPopup =False

  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vTabela As String

  If CurrentQuery.FieldByName("CONTRATODESTINO").IsNull Then
    MsgBox("Informe o contrato de destino")
    cancontinue =False
    Exit Sub
  End If

  Set Interface =CreateBennerObject("Procura.Procurar")

  vColunas ="SAM_MODULO.DESCRICAO"
  vCriterio ="SAM_CONTRATO_MOD.CONTRATO = " +CurrentQuery.FieldByName("CONTRATODESTINO").Value

  vCampos ="Módulo"

  vTabela ="SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]"

  vHandle =Interface.Exec(CurrentSystem,vTabela,vColunas,1,vCampos,vCriterio,"Módulo",True,"")

  If vHandle =0 Then
    CurrentQuery.FieldByName("MODULODESTINO").Value =Null
  Else
    CurrentQuery.FieldByName("MODULODESTINO").Value =vHandle
  End If

  Set Interface =Nothing
End Sub

Public Sub MODULOORIGEM_OnPopup(ShowPopup As Boolean)
  ShowPopup =False

  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vTabela As String

  If CurrentQuery.FieldByName("CONTRATOORIGEM").IsNull Then
    MsgBox("Informe o contrato de origem")
    CanContinue =False
    Exit Sub
  End If

  Set Interface =CreateBennerObject("Procura.Procurar")

  vColunas ="SAM_MODULO.DESCRICAO"
  vCriterio ="SAM_CONTRATO_MOD.CONTRATO = " +CurrentQuery.FieldByName("CONTRATOORIGEM").Value

  vCampos ="Módulo"

  vTabela ="SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]"

  vHandle =Interface.Exec(CurrentSystem,vTabela,vColunas,1,vCampos,vCriterio,"Módulo",True,"")

  If vHandle =0 Then
    CurrentQuery.FieldByName("MODULOORIGEM").Value =Null
  Else
    CurrentQuery.FieldByName("MODULOORIGEM").Value =vHandle
  End If

  Set Interface =Nothing

End Sub

Public Sub TABLE_AfterScroll()
If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      DATAFINAL.ReadOnly =False
Else
DATAFINAL.ReadOnly =True
End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL =NewQuery

  If(Not CurrentQuery.FieldByName("DATAFINAL").IsNull) _
     And(CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
    MsgBox("Data final menor que inicial.")
    CanContinue =False
    Exit Sub
  End If

  SQL.Active =False
  SQL.Clear
  SQL.Add("SELECT *  FROM SAM_TRANSFMOD               ")
  SQL.Add(" WHERE CONTRATOORIGEM = :CONTRATOORIGEM    ")
  SQL.Add("   AND CONTRATODESTINO = :CONTRATODESTINO  ")
  SQL.Add("   AND MODULOORIGEM = :MODULOORIGEM        ")
  SQL.Add("   AND MODULODESTINO = :MODULODESTINO      ")
  SQL.Add("   AND HANDLE <> :HANDLE                   ")
  SQL.Add("   AND (((DATAINICIAL <= :DATAINICIAL) AND (DATAFINAL >= :DATAINICIAL ))  ")
  SQL.Add("    OR ((DATAFINAL >= :DATAFINAL ) AND (DATAINICIAL <= :DATAFINAL ))      ")
  SQL.Add("    OR ((DATAINICIAL >= :DATAINICIAL) AND (DATAFINAL <= :DATAFINAL)))     ")
  SQL.ParamByName("CONTRATOORIGEM").Value =CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger
  SQL.ParamByName("CONTRATODESTINO").Value =CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger
  SQL.ParamByName("MODULOORIGEM").Value =CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
  SQL.ParamByName("MODULODESTINO").Value =CurrentQuery.FieldByName("MODULODESTINO").AsInteger
  SQL.ParamByName("DATAINICIAL").Value =CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  SQL.ParamByName("DATAFINAL").Value =CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  SQL.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active =True

  If Not SQL.EOF Then
    MsgBox("Regra para transferência já existente nesta vigência com mesmo contrato e módulo de origem e destino.")
    CanContinue =False
    Exit Sub
  End If
End Sub

'HASH: 3414443C364F8CEE2D330F73364507C0
 
' MACRO SAM_MIGRACAO_EQCANCADMISSAO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

If CurrentQuery.FieldByName("QTDEMINANO").AsInteger > CurrentQuery.FieldByName("QTDEMAXANO").AsInteger Then
  bsShowMessage("Qtde mínima de anos MAIOR que Qtde máxima de anos!", "E")
  CanContinue = False
  Exit Sub
End If

If CurrentQuery.FieldByName("QTDEMINANO").AsInteger = CurrentQuery.FieldByName("QTDEMAXANO").AsInteger Then
  bsShowMessage("Qtde mínima de anos MAIOR igual a Qtde máxima de anos!", "E")
   CanContinue = False
   Exit Sub
End If

Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Active = False
Consulta.Add("SELECT SUM(CASE WHEN QTDEMINANO = :QTDEMIN THEN       ")
Consulta.Add("               CASE WHEN QTDEMAXANO = :QTDEMAX THEN 1 ")
Consulta.Add("                          ELSE 0 END END) QTDE        ")
Consulta.Add("  FROM SAM_MIGRACAO_EQCANCADMISSAO                    ")
Consulta.Add(" WHERE MIGRACAO = :MIGRACAO                           ")
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    Consulta.Add(" AND HANDLE <> :HANDLE")
    Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
Consulta.ParamByName("QTDEMIN").AsInteger = CurrentQuery.FieldByName("QTDEMINANO").AsInteger
Consulta.ParamByName("QTDEMAX").AsInteger = CurrentQuery.FieldByName("QTDEMAXANO").AsInteger
Consulta.ParamByName("MIGRACAO").AsInteger = CurrentQuery.FieldByName("MIGRACAO").AsInteger
Consulta.Active = True

If Consulta.FieldByName("QTDE").AsInteger > 0 Then
  bsShowMessage("Já existe regra de equivalência com as mesmos anos!", "E")
  CanContinue = False
  Exit Sub
End If


Consulta.Clear
Consulta.Active = False
Consulta.Add("SELECT QTDEMINANO, QTDEMAXANO                                    ")
Consulta.Add("  FROM SAM_MIGRACAO_EQCANCADMISSAO                               ")
Consulta.Add(" WHERE (((:QTDEMIN BETWEEN (QTDEMINANO + 1) AND (QTDEMAXANO-1))  ")
Consulta.Add("    OR   (:QTDEMAX BETWEEN (QTDEMINANO + 1) AND (QTDEMAXANO-1))) ")
Consulta.Add("    OR  (( (QTDEMINANO) BETWEEN :QTDEMIN +1 AND :QTDEMAX - 1  )  ")
Consulta.Add("    OR   ( (QTDEMAXANO) BETWEEN :QTDEMIN +1 AND :QTDEMAX - 1)))  ")
Consulta.Add("   AND MIGRACAO = :MIGRACAO                                      ")
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    Consulta.Add(" AND HANDLE <> :HANDLE")
    Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
Consulta.ParamByName("QTDEMIN").AsInteger = CurrentQuery.FieldByName("QTDEMINANO").AsInteger
Consulta.ParamByName("QTDEMAX").AsInteger = CurrentQuery.FieldByName("QTDEMAXANO").AsInteger
Consulta.ParamByName("MIGRACAO").AsInteger = CurrentQuery.FieldByName("MIGRACAO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("QTDEMINANO").IsNull Or Not Consulta.FieldByName("QTDEMAXANO").IsNull Then
  bsShowMessage("Regra de equivalência informada está cruzando com alguma já existente!", "E")
  CanContinue = False
  Exit Sub
End If

End Sub

'HASH: 2F987AC672ECF8F458A3C04209A52202
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
Dim qWork 
If (CurrentQuery.State = 3) Then    'Se está inserindo uma nova tabela por empresa mestre 
  Set qWork = NewQuery 
  qWork.Add("SELECT NOME,FLAGS FROM Z_TABELAS WHERE HANDLE = " + CurrentQuery.FieldByName("TABELA").AsString) 
  qWork.Active = True 
 
    ' Checar se tabela esta marcada com flag para não utilizar conceito empresa mestre 
    If (qWork.FieldByName("FLAGS").AsInteger And 262144) <> 0 Then   'então está com FLAG checado "Não empresa mestre" 
      Err.Raise vbsUserException, , "Tabela está configurada para não utilizar conceito de empresa mestre." 
    End If 
 
 
  If (qWork.FieldByName("NOME").AsString = "Z_PERIODOS") Then  'e a tabela é a Z_PERIODOS 
    qWork.Active = False 
    qWork.Clear 
    qWork.Add("SELECT COUNT(*) N FROM Z_PERIODOS WHERE " + CompanyField + " IN (SELECT HANDLE FROM " + CompanyTable + " WHERE EMPRESAMESTRE IS NOT NULL)") 
    qWork.Active = True 
    'Verifica se alguma empresa que tem empresa mestre possui períodos cadastrados 
    If qWork.FieldByName("N").AsInteger > 0 Then 
      'Se possui, estes períodos devem ser removidos antes de cadastrar Z_PERIODOS 
      Err.Raise vbsUserException, , "Existem períodos cadastrados para empresas que possuem empresa mestre. Estes períodos devem ser removidos antes de tornar a tabela Z_PERIODOS por empresa mestre." 
    End If 
  End If 
  qWork.Active = False 
  Set qWork = Nothing 
End If 
End Sub 

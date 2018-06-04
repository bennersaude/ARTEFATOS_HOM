'HASH: 6E6BB7B15108326FE21BB230A486E9A4
Option Explicit

Sub main()

  Dim vsQTD_DIAS As String
  Dim vsCaminhoXML As String
  vsCaminhoXML=""
  vsQTD_DIAS = "0"

  If (SessionVar("QTD_DIAS") <>"") Then
	vsQTD_DIAS = SessionVar("QTD_DIAS")
  End If

  If (SessionVar("CAMINHO_XML") <>"") Then
	vsCaminhoXML = SessionVar("CAMINHO_XML")
  End If

  Dim qParametro  As Object
  Set qParametro  = NewQuery

  Dim Iris As Object
  Set Iris = CreateBennerObject("Benner.Saude.Iris.IntegracaoServico.IntegracaoAutorizacaoIris")

  qParametro.Active = False
  qParametro.Add("  SELECT                    ")
  qParametro.Add("         LOCALWSDL,         ")
  qParametro.Add("         USUARIOWS,         ")
  qParametro.Add("         SENHAWS,           ")
  qParametro.Add("         EMPRESAAUDITORIA   ")
  qParametro.Add("    FROM                    ")
  qParametro.Add("         SAM_PARAMETROSWEB  ")
  qParametro.Active = True

  Iris.Url = qParametro.FieldByName("LOCALWSDL").AsString
  Iris.Usuario = qParametro.FieldByName("USUARIOWS").AsString
  Iris.Senha = qParametro.FieldByName("SENHAWS").AsString
  Iris.Empresa = qParametro.FieldByName("EMPRESAAUDITORIA").AsString
  Iris.Company = 1

  qParametro.Active = False
  Set qParametro = Nothing

  Dim msgOut As String
  Dim resultado As Boolean

  resultado = Iris.EnviarXML(ServerDate - CInt(vsQTD_DIAS), ServerDate, vsCaminhoXML, msgOut)

  Set Iris = Nothing

  If (Not resultado) Then
 	  Err.Raise -999,"Erro","Problema na importação de XML: "+ msgOut
  End If

  InfoDescription = "Mensagem de retorno do servidor de auditoria:" + msgOut

End Sub

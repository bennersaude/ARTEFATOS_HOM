'HASH: 762BF42B2915414CFD7D8D176A4FCE34

Public Sub Main


  Dim vDiretorioTemporarioServidor 	    As String
  Dim qParametrosWeb As Object
  Set qParametrosWeb = NewQuery
  qParametrosWeb.Active = False
  qParametrosWeb.Clear
  qParametrosWeb.Add("SELECT DIRETORIOTEMPORARIOSERVIDOR ")
  qParametrosWeb.Add("FROM SAM_PARAMETROSWEB             ")
  qParametrosWeb.Active = True

  vDiretorioTemporarioServidor = qParametrosWeb.FieldByName("DIRETORIOTEMPORARIOSERVIDOR").AsString
  vUltimoCaractere = Mid(vDiretorioTemporarioServidor,Len(vDiretorioTemporarioServidor))
  If vUltimoCaractere <> "\" Then
    vDiretorioTemporarioServidor = vDiretorioTemporarioServidor + "\"
  End If

  If Dir(vDiretorioTemporarioServidor, vbDirectory) <> "" Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Active = False
    SQL.Add("SELECT HANDLE FROM Z_MACROS WHERE NOME LIKE :NOME")
    SQL.ParamByName("NOME").AsString = "GerarDemonstrativo"
    SQL.Active = True

    Dim sx As CSServerExec
    Set sx = NewServerExec
    sx.Description = "Geração dos demonstrativos"
    sx.Process = SQL.FieldByName("HANDLE").AsInteger
    sx.SessionVar("HANDLESOLICIT") = ServiceVar("HANDLESOLICIT")
    sx.SessionVar("TIPO") = ServiceVar("TIPO")
    sx.Execute

    ServiceResult = "Demonstrativo(s) sendo gerado(s) no servidor, aguarde alguns instantes."

    Set sx = Nothing
    Set SQL = Nothing

  Else
	ServiceResult = "Problema com parametrização do diretório temporário (Adm/Parâmetros Gerais/Web - Diretório Temporário do Servidor [" + vDiretorioTemporarioServidor + "]). Favor entrar em contato com a Operadora!"
  End If
  Set qParametrosWeb = Nothing


End Sub


'HASH: F33221C88E854A62F29DD9F805F00E20
Sub Main

  Dim psNotas as String
  Dim psDocumentos as String
  Dim psMensagemRetorno as String

  Dim SQL as Object



  psNotas = ServiceVar("psNotas") 
  psDocumentos = ServiceVar("psDocumentos")

  Dim Interface as Object

  set Interface = createBennerObject("SFNNota.RotinaNota_ConciliarNota")
  Interface.Exec(CurrentSystem, psNotas, psDocumentos, psMensagemRetorno)

 set SQL = newQuery
 SQL.Add("UPDATE ABREV SET TEXTO = :TEXTO WHERE HANDLE = 6")
  SQL.ParamByName("TEXTO").AsString = "teste"
  SQL.ExecSql

  ServiceVar("psMensagemRetorno") =  cstr(psMensagemRetorno)

End Sub

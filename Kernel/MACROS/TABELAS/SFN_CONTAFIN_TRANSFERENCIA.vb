'HASH: B5B23B6A8BBA6EABC8C8A6442CFEC8CA
 
'Macro: SFN_CONTAFIN_TRANSFERENCIA

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("FATURA").IsNull Then
    FATURA.Visible = False
  Else
    FATURA.Visible = True
  End If
  If CurrentQuery.FieldByName("DOCUMENTO").IsNull Then
    DOCUMENTO.Visible = False
  Else
    DOCUMENTO.Visible = True
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT CFIN.TABRESPONSAVEL,")
  SQL.Add("       BEN.BENEFICIARIO,")
  SQL.Add("       BEN.NOME NOMEBENEFICIARIO,")
  SQL.Add("       PES.CNPJCPF,")
  SQL.Add("       PES.NOME NOMEPESSOA")
  SQL.Add("FROM SFN_CONTAFIN CFIN")
  SQL.Add("LEFT JOIN SAM_BENEFICIARIO BEN ON BEN.HANDLE = CFIN.BENEFICIARIO")
  SQL.Add("LEFT JOIN SFN_PESSOA PES ON PES.HANDLE = CFIN.PESSOA")
  SQL.Add("WHERE CFIN.HANDLE = :HCONTAFIN")
  SQL.ParamByName("HCONTAFIN").Value = CurrentQuery.FieldByName("CONTAFINANCEIRAORIGEM").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then 'Beneficário
    ROTULOCONTAFINANCEIRAORIGEM.Text = "Beneficiário responsável de origem: " + SQL.FieldByName("BENEFICIARIO").AsString + " - " + SQL.FieldByName("NOMEBENEFICIARIO").AsString
  Else 'Pessoa
    ROTULOCONTAFINANCEIRAORIGEM.Text = "Pessoa responsável de origem: " + SQL.FieldByName("CNPJCPF").AsString + " - " + SQL.FieldByName("NOMEPESSOA").AsString
  End If

  SQL.Clear
  SQL.Add("SELECT CFIN.TABRESPONSAVEL,")
  SQL.Add("       BEN.BENEFICIARIO,")
  SQL.Add("       BEN.NOME NOMEBENEFICIARIO,")
  SQL.Add("       PES.CNPJCPF,")
  SQL.Add("       PES.NOME NOMEPESSOA")
  SQL.Add("FROM SFN_CONTAFIN CFIN")
  SQL.Add("LEFT JOIN SAM_BENEFICIARIO BEN ON BEN.HANDLE = CFIN.BENEFICIARIO")
  SQL.Add("LEFT JOIN SFN_PESSOA PES ON PES.HANDLE = CFIN.PESSOA")
  SQL.Add("WHERE CFIN.HANDLE = :HCONTAFIN")
  SQL.ParamByName("HCONTAFIN").Value = CurrentQuery.FieldByName("CONTAFINANCEIRADESTINO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then 'Beneficário
    ROTULOCONTAFINANCEIRADESTINO.Text = "Beneficiário responsável de destino: " + SQL.FieldByName("BENEFICIARIO").AsString + " - " + SQL.FieldByName("NOMEBENEFICIARIO").AsString
  Else 'Pessoa
    ROTULOCONTAFINANCEIRADESTINO.Text = "Pessoa responsável de destino: " + SQL.FieldByName("CNPJCPF").AsString + " - " + SQL.FieldByName("NOMEPESSOA").AsString
  End If

  Set SQL = Nothing
End Sub

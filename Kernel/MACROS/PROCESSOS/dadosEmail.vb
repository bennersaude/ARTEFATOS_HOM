'HASH: A4E1DE429A33AA83F0EB8E698640CE4C

Public Sub Main

	Dim qParaProcContas As BPesquisa
	Dim qPrestador As BPesquisa
	Dim qParamProcContas As BPesquisa
	Dim qRegimePagamento As BPesquisa


	Dim vsEmailDestinatario As String
	Dim vsEmailRemetente As String
	Dim vsAssunto As String
	Dim vbTabEnviaEmail As Boolean

	Set qRemetente = NewQuery
	Set qPrestadorEmail = NewQuery
	Set qParamProcContas = NewQuery
	Set qRegimePagamento = NewQuery

	qPrestadorEmail.Add("SELECT P.EMAIL                                 ")
	qPrestadorEmail.Add("  FROM SAM_PRESTADOR P                         ")
	qPrestadorEmail.Add("  JOIN SAM_PEG PG ON (P.HANDLE = PG.RECEBEDOR) ")
	qPrestadorEmail.Add(" WHERE PG.HANDLE = :HANDLE                     ")
	qPrestadorEmail.ParamByName("HANDLE").Value = CInt(ServiceVar("pPeg"))
	qPrestadorEmail.Active = True

	vsEmailDestinatario = qPrestadorEmail.FieldByName("EMAIL").AsString

	qPrestadorEmail.Active = False
	Set qPrestadorEmail = Nothing

	qRemetente.Add("SELECT EMAIL             ")
    qRemetente.Add("  FROM Z_GRUPOUSUARIOS   ")
    qRemetente.Add(" WHERE HANDLE = :HANDLE  ")
    qRemetente.ParamByName("HANDLE").AsInteger = CurrentUser
    qRemetente.Active = True

    vsEmailRemetente = qRemetente.FieldByName("EMAIL").AsString

	qRemetente.Active = False
	Set qRemetente = Nothing

	qParamProcContas.Add("SELECT M.ASSUNTO, P.TABCOMUNICADEVOLUCAO                       ")
	qParamProcContas.Add("  FROM SAM_PARAMETROSPROCCONTAS P                              ")
	qParamProcContas.Add("  JOIN SAM_MENSAGEM_HTML M ON (M.HANDLE = P.MENSAGEMDEVOLUCAO) ")
	qParamProcContas.Active = True

	vsAssunto = qParamProcContas.FieldByName("ASSUNTO").AsString

	qRegimePagamento.Add("SELECT TABREGIMEPGTO    ")
	qRegimePagamento.Add("  FROM SAM_PEG          ")
	qRegimePagamento.Add(" WHERE HANDLE = :HANDLE ")
	qRegimePagamento.ParamByName("HANDLE").Value = CInt(ServiceVar("pPeg"))
	qRegimePagamento.Active = True

	If ((qParamProcContas.FieldByName("TABCOMUNICADEVOLUCAO").AsInteger = 1) And _
	    (qRegimePagamento.FieldByName("TABREGIMEPGTO").AsInteger) = 1) Then

	  vbTabEnviaEmail = True
	Else
	  vbTabEnviaEmail = False
	End If

	qParamProcContas.Active = False
	Set qParamProcContas = Nothing

	qRegimePagamento.Active = False
	Set qRegimePagamento = Nothing

    ServiceVar("pDestinatario") = vsEmailDestinatario
    ServiceVar("pRemetente") = vsEmailRemetente
    ServiceVar("pAssunto") = vsAssunto

    ServiceResult = vbTabEnviaEmail

	End Sub

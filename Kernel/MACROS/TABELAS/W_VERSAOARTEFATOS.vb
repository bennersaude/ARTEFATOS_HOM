'HASH: A95196ABCB514FA9865F0473F68E1650
 
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean) 
 
If CommandID = "StartInstallArtifacts" Then 
 
  Dim cx As CSServerExec 
  Set cx = NewServerExec 
  cx.DllClassName = "Benner.Tecnologia.Wes.Metadata.InstallArtifactsProcessHelper" 
  cx.Description = "Processo de instalação de artefatos do WES" 
  cx.Execute 
  InfoDescription = CStr(cx.ProcessLog.Handle) 
  Set cx = Nothing 
 
End If 
End Sub 
 

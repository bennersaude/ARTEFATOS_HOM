'HASH: 093C3E8C64DDE0E61C55180DAC87E276

Option Explicit

Public Sub Main
  Dim dll As Object
  Dim msg As String

  Set dll = CreateBennerObject("sampegdigit.digitacao")
  dll.Exec(CurrentSystem, 0, msg)

  Set dll = Nothing
End Sub

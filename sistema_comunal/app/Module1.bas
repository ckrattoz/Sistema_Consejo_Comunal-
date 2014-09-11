Attribute VB_Name = "Module1"
Public cnn As ADODB.Connection
Public Sub getConnected()
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\usuarios.mdb" & "; Persist Security Info=False;"
cnn.Open
End Sub


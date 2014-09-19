Attribute VB_Name = "Module1"
Public cnn As ADODB.Connection
Public Sub getConnected()
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\usuarios.mdb" & ";Jet OLEDB:Database Password=sccadm1n; Persist Security Info=False;"
cnn.Open
End Sub

Public Function getDay() As String
    getDay = Day(Now)
End Function

Public Function getMonth() As String
    mesNumero = Month(Now)
    Select Case mesNumero
        Case "1"
             mes = "ENERO"
        Case "2"
             mes = "FEBRERO"
        Case "3"
             mes = "MARZO"
        Case "4"
             mes = "ABRIL"
        Case "5"
             mes = "MAYO"
        Case "6"
             mes = "JUNIO"
        Case "7"
             mes = "JULIO"
        Case "8"
             mes = "AGOSTO"
        Case "9"
             mes = "SEPTIEMBRE"
        Case "10"
             mes = "OCTUBRE"
        Case "11"
             mes = "NOVIEMBRE"
        Case "12"
             mes = "DICIEMBRE"
        Case Else
            mes = " "
    End Select
    getMonth = mes
End Function
Public Function getYear() As String
    fullYear = Format(Now, "yy")
    Select Case fullYear
        Case "13"
             ano = "TRECE"
        Case "14"
             ano = "CATORCE"
        Case "15"
             ano = "QUINCE"
        Case "16"
             ano = "DIECISEIS"
        Case "17"
             ano = "DIECISIETE"
        Case "18"
             ano = "DIECIOCHO"
        Case "19"
             ano = "DIECINUEVE"
        Case "20"
             ano = "VEINTE"
        Case "21"
             ano = "VEINTIUNO"
        Case "22"
             ano = "VEINTIDOS"
        Case "23"
             ano = "VEINTITRES"
        Case "24"
             ano = "VEINTICUATRO"
        Case "25"
             ano = "VEINTICINCO"
        Case Else
            ano = "mmmmmmmmmm "
    End Select
    getYear = ano
End Function


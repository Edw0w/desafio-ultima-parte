Dim cpf As String
cpf = Sheets("CADASTRO-CLIENTE").Range("C17").Value

' Verifica se contém apenas números ou está vazio
If cpf Like "*[!0-9]*" Or cpf = "" Then
    MsgBox "Digite apenas números no campo CPF.", vbExclamation
    Sheets("CADASTRO-CLIENTE").Range("C17").ClearContents
    Exit Sub
End If

' Verifica se tem exatamente 11 dígitos
If Len(cpf) <> 11 Then
    MsgBox "O CPF deve conter exatamente 11 números.", vbExclamation
    Sheets("CADASTRO-CLIENTE").Range("C17").ClearContents
    Exit Sub
End If

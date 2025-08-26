VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SalvarBt_Click()

Dim cnt As Control

For Each cnt In Me.Controls
    If TypeOf cnt Is MSForms.TextBox Then
        If cnt.Value = "" Then
            MsgBox "O campo " & cnt.Name & " está vazio. Por favor preencha corretamente!", vbCritical
            cnt.SetFocus
            Exit Sub
        End If
    End If
Next cnt



Call CadastraNovoUser


End Sub


Sub CadastraNovoUser()

Call ConexaoBd
         
Sql = "SELECT COUNT(*) FROM Usuarios WHERE Login = '" & Me.LoginTx.Text & "'"

'------------------------------------------'
Sql = "SELECT * FROM Usuarios"

Rs.Open Sql, cnn, adOpenKeyset, adLockOptimistic

Rs.AddNew

    Rs!Login = Me.LoginTx.Value
    Rs!Nome = Me.NomeTx.Value
    Rs!Sobrenome = Me.SobrenomeTx.Value
    Rs!Email = Me.EmailTx.Value
    Rs!Senha = Me.SenhaTx.Value
    Rs!UserMaster = Me.UserMasterSimOp.Value
    Rs!DtCadastro = Now
    Rs!Status = "Habilitado"
    
    Rs.Update

Rs.Close
cnn.Close

Set Rs = Nothing
Set cnn = Nothing

MsgBox "Operação realizada com sucesso", vbInformation, "Cadastrado"

End Sub

Private Sub UserForm_Click()

End Sub

# myAHKv1
Meus códigos para todos.

# DLL
## Netapi32.dll
### Criar Usuário (NetUserAdd.ahk)
Cria um novo usuário no sistema.
```ahk
username := "Teste"
password := "teste1234"
description := "Teste de usuário."

CreateUser(username, password, description)
```

### Alterar Senha do Usuário (NetUserChangePassword.ahk)
Altera a senha de um usuário utilizando a senha antiga.

```ahk
username := "Teste"
oldPassword := "teste1234"
newPassword := "1234teste"
domain := ""

ChangeUserPasswordWithOldPassword(domain, username, oldPassword, newPassword)
```

### Alterar Informações do Usuário (NetUserInfo.ahk)
Atualiza informações de um usuário, como senha, descrição, ativação de conta e nome completo.

```ahk
username := "Teste"
fullname := "Teste Completo"
password := "teste1234"
description := "Teste de usuário alterando informações."

;Alterando a Senha
ChangeUserPassword(username, password)

;Alterando descrição
SetUserInfoDescription(username, description)

;Ativando conta do Usuário
ActivateUserAccount(username)

;Alterando Nome Completo do Usuário
SetUserInfoFullName(username, fullname)
```

### Adicionar Usuário ao Grupo Administrador(NetLocalGroupAddMembres.ahk)
Adiciona um usuário ao grupo de administradores locais.

```ahk
username := "Teste"

AddToAdministrators(username)
```

### Listar Usuário do Grupo Administrador(NetLocalGroupGetMembres.ahk)
Lista os membros do grupo de administradores locais.
```ahk
GetLocalGroupMembers()
```

### Excluir Usuário (NetUserDel.ahk)
Remove um usuário do sistema.
```ahk
username := "Teste"

DeleteUser(username)
```
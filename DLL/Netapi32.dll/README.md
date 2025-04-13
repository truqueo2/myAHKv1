# Netapi32.dll
## NetLocalGroupAddMembers.ahk

### Descrição
Este script utiliza a função `NetLocalGroupAddMembers` da DLL `Netapi32.dll` para adicionar um usuário ao grupo de administradores locais no Windows. Ele verifica se o script está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `AddToAdministrators(username)`
Adiciona um usuário ao grupo de administradores locais.

- **Parâmetros:**
  - `username`: Nome do usuário que será adicionado ao grupo de administradores.

- **Funcionamento:**
  - Cria um buffer para armazenar informações do grupo.
  - Converte o nome do usuário para Unicode, se necessário.
  - Chama a função `NetLocalGroupAddMembers` da DLL `Netapi32.dll` para adicionar o usuário ao grupo.

#### 2. `AnsiToUnicode(str)`
Converte uma string ANSI para Unicode.

- **Parâmetros:**
  - `str`: String em ANSI que será convertida.

- **Retorno:**
  - String convertida para Unicode.

### Exemplo de Uso

```ahk
; Exemplo de como usar o script para adicionar um usuário ao grupo de administradores
username := "UsuarioTeste" ; Substitua pelo nome do usuário desejado
AddToAdministrators(username)

MsgBox, 64, Sucesso, O usuário %username% foi adicionado ao grupo de administradores locais!
```
## NetLocalGroupGetMembers.ahk

### Descrição
Este script utiliza a função `NetLocalGroupGetMembers` da DLL `Netapi32.dll` para listar os membros de um grupo local no Windows. Ele verifica se o script está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `GetLocalGroupMembers()`
Obtém os membros de um grupo local.

- **Funcionamento:**
  - Chama a função `NetLocalGroupGetMembers` da DLL `Netapi32.dll` para obter os membros do grupo local especificado.
  - Processa os dados retornados e exibe os nomes dos membros.
  - Libera a memória alocada pela função após o uso.

#### 2. `AnsiToUnicode(str)`
Converte uma string ANSI para Unicode.

- **Parâmetros:**
  - `str`: String em ANSI que será convertida.

- **Retorno:**
  - String convertida para Unicode.

### Exemplo de Uso

```ahk
; Exemplo de como usar o script para listar os membros do grupo "Administradores"
GetLocalGroupMembers()
```

## NetUserAdd.ahk

### Descrição
Este script utiliza a função `NetUserAdd` da DLL `Netapi32.dll` para criar um novo usuário no Windows. Ele permite especificar o nome do usuário, senha e uma descrição. O script verifica se está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `CreateUser(username, password, description)`
Cria um novo usuário no sistema.

- **Parâmetros:**
  - `username`: Nome do usuário a ser criado.
  - `password`: Senha do usuário.
  - `description`: Descrição do usuário.

- **Funcionamento:**
  - Preenche a estrutura `USER_INFO_1` com os dados fornecidos.
  - Chama a função `NetUserAdd` da DLL `Netapi32.dll` para criar o usuário.
  - Exibe uma mensagem de sucesso ou erro com base no resultado da operação.

#### 2. `AnsiToUnicode(str)`
Converte uma string ANSI para Unicode.

- **Parâmetros:**
  - `str`: String em ANSI que será convertida.

- **Retorno:**
  - String convertida para Unicode.

### Exemplo de Uso

```ahk
; Exemplo de como usar o script para criar um novo usuário
username := "NovoUsuario" ; Substitua pelo nome do usuário desejado
password := "Senha123"    ; Substitua pela senha desejada
description := "Usuário criado via script" ; Substitua pela descrição desejada

resultado := CreateUser(username, password, description)

MsgBox, %resultado%
```

## NetUserChangePassword.ahk

### Descrição
Este script utiliza a função `NetUserChangePassword` da DLL `Netapi32.dll` para alterar a senha de um usuário no Windows. Ele permite especificar o domínio (ou computador local), o nome do usuário, a senha antiga e a nova senha. O script verifica se está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `ChangeUserPasswordWithOldPassword(domain, username, oldPassword, newPassword)`
Altera a senha de um usuário local ou de domínio.

- **Parâmetros:**
  - `domain`: Nome do domínio (ou `NULL` para o computador local).
  - `username`: Nome do usuário cuja senha será alterada.
  - `oldPassword`: Senha antiga do usuário.
  - `newPassword`: Nova senha do usuário.

- **Funcionamento:**
  - Chama a função `NetUserChangePassword` da DLL `Netapi32.dll` para alterar a senha do usuário.
  - Exibe uma mensagem de sucesso ou erro com base no resultado da operação.

### Exemplo de Uso

```ahk
; Exemplo de como usar o script para alterar a senha de um usuário
domain := ""               ; Deixe vazio para o computador local
username := "UsuarioTeste" ; Substitua pelo nome do usuário
oldPassword := "SenhaAntiga123" ; Substitua pela senha antiga
newPassword := "NovaSenha456"   ; Substitua pela nova senha

ChangeUserPasswordWithOldPassword(domain, username, oldPassword, newPassword)
```

## NetUserDel.ahk

### Descrição
Este script utiliza a função `NetUserDel` da DLL `Netapi32.dll` para excluir um usuário no Windows. Ele permite especificar o nome do usuário a ser removido do sistema. O script verifica se está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `DeleteUser(username)`
Exclui um usuário local ou de domínio.

- **Parâmetros:**
  - `username`: Nome do usuário a ser excluído.

- **Funcionamento:**
  - Chama a função `NetUserDel` da DLL `Netapi32.dll` para excluir o usuário especificado.
  - Exibe uma mensagem de sucesso ou erro com base no resultado da operação.

### Exemplo de Uso

```ahk
; Exemplo de como usar o script para excluir um usuário
username := "UsuarioTeste" ; Substitua pelo nome do usuário que deseja excluir

DeleteUser(username)
```

## NetUserSetInfo.ahk

### Descrição
Este script utiliza a função `NetUserSetInfo` da DLL `Netapi32.dll` para atualizar informações de um usuário no Windows. Ele permite alterar a senha, descrição, ativar/desativar contas e atualizar o nome completo do usuário. O script verifica se está sendo executado como administrador e, caso contrário, tenta reiniciar com privilégios elevados.

### Funções

#### 1. `ChangeUserPassword(username, password)`
Atualiza a senha de um usuário.

- **Parâmetros:**
  - `username`: Nome do usuário.
  - `password`: Nova senha do usuário.

#### 2. `SetUserInfoDescription(username, description)`
Atualiza a descrição de um usuário.

- **Parâmetros:**
  - `username`: Nome do usuário.
  - `description`: Nova descrição associada à conta do usuário.

#### 3. `ActivateUserAccount(username)`
Ativa a conta de um usuário.

- **Parâmetros:**
  - `username`: Nome do usuário cuja conta será ativada.

#### 4. `SetUserInfoFullName(username, fullname)`
Atualiza o nome completo de um usuário.

- **Parâmetros:**
  - `username`: Nome do usuário.
  - `fullname`: Novo nome completo do usuário.

### Exemplos de Uso

#### Alterar a senha de um usuário
```ahk
username := "UsuarioTeste" ; Substitua pelo nome do usuário
password := "NovaSenha123" ; Substitua pela nova senha

ChangeUserPassword(username, password)
MsgBox, Senha alterada com sucesso!
```
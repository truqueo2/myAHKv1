/*
DLL: Netapi32.dll
Autor: truqueo2
Link: https://learn.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netuserchangepassword

Estrutura NetUserChangePassword
NET_API_STATUS NET_API_FUNCTION NetUserChangePassword(
  [in] LPCWSTR domainname,  ; Nome do domínio (NULL para o computador local)
  [in] LPCWSTR username,    ; Nome do usuário
  [in] LPCWSTR oldpassword, ; Senha antiga
  [in] LPCWSTR newpassword  ; Nova senha
);
*/

; Verifica se o script está sendo executado como administrador
if !A_IsAdmin {
  try {
    Run *RunAs "%A_ScriptFullPath%" ; Tenta executar o script como administrador
    ExitApp
  } catch e {
    MsgBox, 16, Erro, Não foi possível Executar como Administrador.
    ExitApp
  }
}

; Função para alterar a senha de um usuário local ou de domínio
ChangeUserPasswordWithOldPassword(domain, username, oldPassword, newPassword) {
    ; Se o domínio não for especificado, define como NULL
    domain := domain ? domain : 0

    ; Chamando a função NetUserChangePassword
    result := DllCall("Netapi32.dll\NetUserChangePassword"
        , "Ptr", domain          ; Nome do domínio ou NULL para o computador local
        , "Ptr", &username        ; Nome do usuário
        , "Ptr", &oldPassword     ; Senha antiga
        , "Ptr", &newPassword     ; Nova senha
        , "UInt")                 ; Tipo de retorno (NET_API_STATUS)

    ; Verifica o resultado da chamada
    if (result != 0) {
        ; Exibe mensagem de erro com o código retornado
        MsgBox, Erro ao alterar a senha do usuário (código %result%)
    } else {
        ; Exibe mensagem de sucesso
        MsgBox, Senha do usuário alterada com sucesso!
    }
}
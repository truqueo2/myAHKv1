/*
DLL: Netapi32.dll
Autor: truqueo2
Link: https://learn.microsoft.com/pt-br/windows/win32/api/lmaccess/nf-lmaccess-netusersetinfo

Estrutura NetUserSetInfo
NET_API_STATUS NET_API_FUNCTION NetUserSetInfo(
  [in]  LPCWSTR servername, ; Nome do servidor (NULL para o servidor local)
  [in]  LPCWSTR username,   ; Nome do usuário
  [in]  DWORD   level,      ; Nível de informação
  [in]  LPBYTE  buf,        ; Ponteiro para a estrutura de informações do usuário
  [out] LPDWORD parm_err    ; Parâmetro com erro (NULL se não for necessário)
);

Level:
1003 - Especifica uma senha de usuário.
1007 - Especifica um comentário a ser associado à conta de usuário.
1008 - Especifica atributos de conta de usuário (Ativar a conta).
1011 - Especifica o nome completo do usuário.
... Consulte o link para outros níveis.
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

ChangeUserPassword(username, password) {
  ; Atualizando a senha do usuário
  VarSetCapacity(user_info, 8 + (StrLen(password) + 1) * 2, 0) ; Aloca memória para a estrutura
  passwordPtr := A_IsUnicode ? &password : AnsiToUnicode(password) ; Converte a senha para Unicode
  NumPut(passwordPtr, user_info, 0, "Ptr") ; Preenche a estrutura com a nova senha

  ; Chama a função NetUserSetInfo para atualizar a senha
  DllCall("Netapi32.dll\NetUserSetInfo"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "Str", username      ; Nome do usuário
    , "UInt", 1003         ; Nível 1003: Atualiza a senha
    , "Ptr", &user_info    ; Ponteiro para a estrutura de informações
    , "UInt*", 0)          ; Parâmetro com erro (não usado aqui)
}

SetUserInfoDescription(username, description) {
  ; Atualizando a descrição do usuário
  VarSetCapacity(user_info, 8 + (StrLen(description) + 1) * 2, 0) ; Aloca memória para a estrutura
  descriptionPtr := A_IsUnicode ? &description : AnsiToUnicode(description) ; Converte a descrição para Unicode
  NumPut(descriptionPtr, user_info, 0, "Ptr") ; Preenche a estrutura com a nova descrição

  ; Chama a função NetUserSetInfo para atualizar a descrição
  DllCall("Netapi32.dll\NetUserSetInfo"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "Str", username      ; Nome do usuário
    , "UInt", 1007         ; Nível 1007: Atualiza o comentário (descrição)
    , "Ptr", &user_info    ; Ponteiro para a estrutura de informações
    , "UInt*", 0)          ; Parâmetro com erro (não usado aqui)
}

ActivateUserAccount(username) {
  ; Ativando a conta do usuário
  flags := 0x0001  ; UF_SCRIPT: Remove a flag de conta desativada (UF_ACCOUNTDISABLE)
  VarSetCapacity(user_info, 4, 0) ; Aloca memória para a estrutura
  NumPut(flags, user_info, 0, "UInt") ; Preenche a estrutura com a flag para ativar a conta

  ; Chama a função NetUserSetInfo para ativar a conta
  result := DllCall("Netapi32.dll\NetUserSetInfo"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "Str", username      ; Nome do usuário
    , "UInt", 1008         ; Nível 1008: Atualiza atributos da conta
    , "Ptr", &user_info    ; Ponteiro para a estrutura de informações
    , "UInt*", 0)          ; Parâmetro com erro (não usado aqui)

  ; Verifica o resultado da chamada
  if (result != 0) {
    MsgBox, Erro ao ativar a conta do usuário (código %result%)
  } else {
    MsgBox, Conta do usuário ativada com sucesso!
  }
}

SetUserInfoFullName(username, fullname) {
  ; Atualizando o nome completo do usuário
  VarSetCapacity(user_info, 8 + (StrLen(fullname) + 1) * 2, 0) ; Aloca memória para a estrutura
  fullnamePtr := A_IsUnicode ? &fullname : AnsiToUnicode(fullname) ; Converte o nome completo para Unicode
  NumPut(fullnamePtr, user_info, 0, "Ptr") ; Preenche a estrutura com o novo nome completo

  ; Chama a função NetUserSetInfo para atualizar o nome completo
  DllCall("Netapi32.dll\NetUserSetInfo"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "Str", username      ; Nome do usuário
    , "UInt", 1011         ; Nível 1011: Atualiza o nome completo
    , "Ptr", &user_info    ; Ponteiro para a estrutura de informações
    , "UInt*", 0)          ; Parâmetro com erro (não usado aqui)
}

AnsiToUnicode(str) {
  ; Função para converter uma string ANSI para Unicode
  VarSetCapacity(wstr, (StrLen(str) + 1) * 2, 0) ; Aloca memória para a string Unicode
  DllCall("MultiByteToWideChar"
    , "UInt", 0           ; Código de página (CP_ACP)
    , "UInt", 0           ; Flags (nenhum)
    , "Str", str          ; String ANSI de entrada
    , "Int", -1           ; Tamanho da string de entrada (-1 para calcular automaticamente)
    , "Ptr", &wstr        ; Ponteiro para a string Unicode de saída
    , "Int", StrLen(str) + 1) ; Tamanho do buffer de saída
  Return wstr               ; Retorna a string Unicode
}
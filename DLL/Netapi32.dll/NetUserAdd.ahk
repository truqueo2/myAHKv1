/*
DLL: Netapi32.dll
Autor: truqueo2
Links:
  - https://learn.microsoft.com/pt-br/windows/win32/api/lmaccess/nf-lmaccess-netuseradd
  - https://learn.microsoft.com/pt-br/windows/win32/api/lmaccess/ns-lmaccess-user_info_1

Estrutura NetUserAdd
NET_API_STATUS NET_API_FUNCTION NetUserAdd(
  [in]  LPCWSTR servername, ; Nome do servidor (NULL para o servidor local)
  [in]  DWORD   level,      ; Nível de informação (1 neste caso)
  [in]  LPBYTE  buf,        ; Ponteiro para a estrutura USER_INFO_1
  [out] LPDWORD parm_err    ; Parâmetro com erro (NULL se não for necessário)
);

Estrutura USER_INFO_1
typedef struct _USER_INFO_1 {
  LPWSTR usri1_name;         ; Nome do usuário
  LPWSTR usri1_password;     ; Senha do usuário
  DWORD  usri1_password_age; ; Idade da senha (não usado aqui)
  DWORD  usri1_priv;         ; Nível de privilégio (1 = USER_PRIV_USER)
  LPWSTR usri1_home_dir;     ; Diretório inicial (não usado aqui)
  LPWSTR usri1_comment;      ; Descrição do usuário
  DWORD  usri1_flags;        ; Flags (ex.: UF_SCRIPT | UF_NORMAL_ACCOUNT)
  LPWSTR usri1_script_path;  ; Caminho do script de login (não usado aqui)
} USER_INFO_1, *PUSER_INFO_1, *LPUSER_INFO_1;
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

CreateUser(username, password, description) {
  ; Calcula o tamanho da estrutura USER_INFO_1 com base no tamanho do ponteiro (32 ou 64 bits)
  structSize := A_PtrSize = 8 ? 72 : 40
  VarSetCapacity(user_info, structSize, 0) ; Aloca memória para a estrutura

  ; Converte os parâmetros para Unicode (LPWSTR)
  usernamePtr := A_IsUnicode ? &username : AnsiToUnicode(username)
  passwordPtr := A_IsUnicode ? &password : AnsiToUnicode(password)
  descriptionPtr := A_IsUnicode ? &description : AnsiToUnicode(description)

  ; Preenche a estrutura USER_INFO_1
  NumPut(usernamePtr, user_info, 0, "Ptr")                    ; Nome do usuário (usri1_name)
  NumPut(passwordPtr, user_info, A_PtrSize, "Ptr")            ; Senha do usuário (usri1_password)
  NumPut(0, user_info, A_PtrSize * 2, "UInt")                 ; Idade da senha (usri1_password_age)
  NumPut(1, user_info, A_PtrSize * 2 + 4, "UInt")             ; Privilégio (usri1_priv = USER_PRIV_USER)
  NumPut(0, user_info, A_PtrSize * 2 + 4 * 2, "Ptr")          ; Diretório inicial (usri1_home_dir)
  NumPut(descriptionPtr, user_info, A_PtrSize * 3 + 4 * 2, "Ptr") ; Descrição (usri1_comment)
  NumPut(0x0201, user_info, A_PtrSize * 4 + 4 * 2, "UInt")    ; Flags (UF_SCRIPT | UF_NORMAL_ACCOUNT)
  NumPut(0, user_info, A_PtrSize * 4 + 4 * 3, "Ptr")          ; Caminho do script de login (usri1_script_path)

  ; Chama a função NetUserAdd para criar o usuário
  result := DllCall("Netapi32.dll\NetUserAdd"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "UInt", 1            ; Nível de informação (1)
    , "Ptr", &user_info    ; Ponteiro para a estrutura USER_INFO_1
    , "UInt*", 0)          ; Parâmetro com erro (não usado aqui)

  ; Verifica o resultado da chamada
  if (result != 0) {
    ; Exibe mensagem de erro com o código retornado
    MsgBox, Erro ao criar usuário (código %result%)
    return "Erro ao criar usuário (código " . result . ")"
  }

  ; Exibe mensagem de sucesso
  MsgBox, Usuário criado com sucesso!
  return "Usuário criado com sucesso!"
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
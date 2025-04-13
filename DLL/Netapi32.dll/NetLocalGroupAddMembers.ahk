/*
DLL: Netapi32.dll
Autor: truqueo2
Link: https://learn.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netlocalgroupaddmembers

Estrutura NetLocalGroupAddMembers
NET_API_STATUS NET_API_FUNCTION NetLocalGroupAddMembers(
  [in] LPCWSTR servername,   ; Nome do servidor (NULL para o servidor local)
  [in] LPCWSTR groupname,    ; Nome do grupo local
  [in] DWORD   level,        ; Nível de informação (3 neste caso)
  [in] LPBYTE  buf,          ; Ponteiro para a estrutura de informações do grupo
  [in] DWORD   totalentries  ; Número total de entradas
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

AddToAdministrators(username) {
  ; Função para adicionar um usuário ao grupo de administradores locais
  group := "Administradores" ; Nome do grupo local (em português)

  ; Cria um buffer para armazenar informações do grupo
  VarSetCapacity(group_info, 8, 0)

  ; Converte o nome de usuário para Unicode, se necessário
  usernamePtr := A_IsUnicode ? &username : AnsiToUnicode(username)

  ; Armazena o ponteiro do nome de usuário na estrutura group_info
  NumPut(usernamePtr, group_info, 0, "Ptr")

  ; Chama a função NetLocalGroupAddMembers da DLL Netapi32.dll
  result := DllCall("Netapi32.dll\NetLocalGroupAddMembers"
    , "Ptr", 0             ; Nome do servidor (NULL para o servidor local)
    , "Str", group         ; Nome do grupo local
    , "UInt", 3            ; Nível de informação (3)
    , "Ptr", &group_info   ; Ponteiro para a estrutura de informações do grupo
    , "UInt", 1)           ; Número total de entradas (1 neste caso)
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
  return wstr               ; Retorna a string Unicode
}

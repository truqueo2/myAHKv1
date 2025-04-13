/*
DLL: Netapi32.dll
Autor: truqueo2
Link: https://learn.microsoft.com/pt-br/windows/win32/api/lmaccess/nf-lmaccess-netlocalgroupgetmembers

Estrutura NetLocalGroupGetMembers
NET_API_STATUS NET_API_FUNCTION NetLocalGroupGetMembers(
  [in]      LPCWSTR    servername,       ; Nome do servidor (NULL para o computador local)
  [in]      LPCWSTR    localgroupname,   ; Nome do grupo local
  [in]      DWORD      level,            ; Nível de informação
  [out]     LPBYTE     *bufptr,          ; Ponteiro para o buffer de saída
  [in]      DWORD      prefmaxlen,       ; Tamanho máximo preferido do buffer
  [out]     LPDWORD    entriesread,      ; Número de entradas lidas
  [out]     LPDWORD    totalentries,     ; Número total de entradas disponíveis
  [in, out] PDWORD_PTR resumehandle      ; Ponteiro para continuar listagem (NULL para começar do início)
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

GetLocalGroupMembers(){
  ; Inicializa as variáveis para armazenar os resultados da função
  pBuf := 0               ; Ponteiro para o buffer de saída
  entriesRead := 0        ; Número de entradas lidas
  totalEntries := 0       ; Número total de entradas disponíveis

  ; Chamada da função NetLocalGroupGetMembers da DLL Netapi32.dll
  result := DllCall("Netapi32.dll\NetLocalGroupGetMembers"
    , "ptr", 0             ; Nome do servidor (NULL para o computador local)
    , "str", "Administradores" ; Nome do grupo local
    , "UInt", 3            ; Nível de informação (3 para obter nomes de membros)
    , "ptr*", pBuf         ; Ponteiro para o buffer de saída
    , "UInt", -1           ; Tamanho máximo preferido (-1 para ilimitado)
    , "UInt*", entriesRead ; Número de entradas lidas
    , "UInt*", totalEntries ; Número total de entradas disponíveis
    , "ptr", 0)            ; Ponteiro para continuar listagem (NULL para começar do início)

  ; Verifica o resultado da chamada para identificar erros
  if (result != 0) {
    MsgBox, Erro ao chamar NetLocalGroupGetMembers. Código de erro: %result%
    return
  }

  ; Processa os dados retornados se houver entradas lidas
  if (entriesRead > 0) {
    ; Cada entrada é uma estrutura LOCALGROUP_MEMBERS_INFO_3
    ; O tamanho da estrutura é igual ao tamanho de um ponteiro
    structSize := A_PtrSize

    ; Itera sobre as entradas retornadas
    Loop, %entriesRead% {
      offset := (A_Index - 1) * structSize ; Calcula o deslocamento para cada entrada
      memberPtr := NumGet(pBuf + offset, 0, "ptr") ; Obtém o ponteiro para o nome do membro
      if (memberPtr) {
        memberName := StrGet(memberPtr, "UTF-16") ; Obtém o nome do membro como string UTF-16
        userName := StrSplit(memberName, "\").2 ; Divide pelo '\' e pega apenas o nome de usuário
        MsgBox, Membro %A_Index%: %userName% ; Exibe o nome do membro
      } else {
        MsgBox, Membro %A_Index%: Nome não encontrado. ; Caso o nome não seja encontrado
      }
    }
  }
  ; Libera a memória alocada pelo NetLocalGroupGetMembers
  DllCall("Netapi32.dll\NetApiBufferFree", "ptr", pBuf)
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
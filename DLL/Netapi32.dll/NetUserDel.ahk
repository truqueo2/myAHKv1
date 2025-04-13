/*
DLL: Netapi32.dll
Autor: truqueo2
Link: https://learn.microsoft.com/pt-br/windows/win32/api/lmaccess/nf-lmaccess-netuserdel

Estrutura NetUserDel
NET_API_STATUS NET_API_FUNCTION NetUserDel(
  [in] LPCWSTR servername, ; Nome do servidor (NULL para o servidor local)
  [in] LPCWSTR username    ; Nome do usuário a ser excluído
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

DeleteUser(username) {
  ; Função para excluir um usuário local ou de domínio
  ; Chamando a função NetUserDel da DLL Netapi32.dll
  result := DllCall("Netapi32.dll\NetUserDel"
        , "Ptr", 0         ; Nome do servidor (NULL para o servidor local)
        , "Str", username) ; Nome do usuário a ser excluído

  ; Verifica o resultado da chamada
  if (result != 0) {
    ; Exibe mensagem de erro com o código retornado
    MsgBox, Erro ao excluir o usuário (código %result%)
  } else {
    ; Exibe mensagem de sucesso
    MsgBox, Usuário excluído com sucesso!
  }
}

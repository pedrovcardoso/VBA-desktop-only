# VBA-desktop-only

Permite edição de planilhas Excel apenas no aplicativo desktop, protegendo contra edição no Excel Online.  
O código usa VBA para proteger/desproteger células e interceptar atalhos de teclado.

### Como funciona
- Todas as planilhas são protegidas automaticamente na abertura.
- Apenas a célula ativa pode ser editada.
- Intercepta:
  - Ctrl+C → cópia controlada
  - Ctrl+V → colagem controlada
  - Tab e Shift+Tab → movimentação personalizada
- Funciona somente no Excel Desktop. No Excel Online a planilha permanece bloqueada.

### Instalação
1. Abra o Editor VBA no Excel (Alt+F11).
2. Importe o arquivo de eventos `EstaPastaDeTrabalho.cls` no projeto.
3. Importe o módulo auxiliar `Módulo1.bas` no projeto.
4. Altere a senha padrão na constante `minhaSenha` do Módulo1.
5. Salve o arquivo como `.xlsm`.
6. Feche e reabra a planilha para ativar o código.
7. Ao abrir a planilha, pode aparecer o aviso de segurança "As macros foram desabilitadas" (barra amarela no topo do Excel).  
   Clique em "Habilitar Conteúdo" para que o código VBA funcione corretamente.


### Limitações
- Copiar e colar leva apenas o texto, sem estilos ou formatação.
- Não permite mover a célula ou intervalo (arrastar e soltar).
- Está bloqueado apenas para edição no Excel Online.  
  - Ainda é possível colorir ou formatar a célula, mas isso pode ser alterado no código.

### Observações
- A senha fica visível no código VBA. Não use senhas sensíveis.
- Pode haver impacto de desempenho em planilhas muito grandes.
- Compatível com Excel 2007 ou superior, 32 e 64 bits.
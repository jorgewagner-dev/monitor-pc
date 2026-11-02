# Monitor PC

Painel para Windows 10/11 que acompanha CPU, RAM, disco, placa de vídeo e rede em tempo real, com alertas,
FPS dos jogos, diagnóstico de quedas, Modo Jogo automático, verificação de segurança, vida útil das peças
e comandos de otimização. Feito em PowerShell + WPF, sem instalar nada.

Dois perfis, escolhidos na primeira vez que abre (e trocados depois na Aparência):

- **Gamer** — FPS ao vivo, quedas e causas, Modo Jogo, ping e sessões de jogo.
- **Profissional** — Modo Reunião (protege chamadas detectando o microfone em uso), inventário e conformidade
  (licença, Windows Update, BitLocker, Secure Boot, TPM, backup) e relatório do PC em HTML/PDF para o TI.

Manual completo: [Manual do Monitor PC.docx](Manual%20do%20Monitor%20PC.docx).

## Como usar

1. Baixe o `MonitorPC.zip` ou os arquivos `MonitorPC.ps1` e `Iniciar Monitor PC.bat` para uma pasta.
2. Dê dois cliques em `Iniciar Monitor PC.bat`.
3. Se o Windows avisar que protegeu o computador: **Mais informações → Executar assim mesmo**.

Detalhes em [LEIA-ME.txt](LEIA-ME.txt).

## Atualizações

O Monitor confere uma vez por dia o arquivo [`versao.json`](versao.json). Quando há versão nova, ele baixa o
`MonitorPC.ps1`, confere a impressão digital SHA-256 e só então instala (mantendo a versão anterior como `.bak`).

Para publicar uma versão nova: `.\publicar-versao.ps1 -Versao 2.6.0 -Notas "o que mudou"`.

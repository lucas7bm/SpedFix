# SpedFix
Este script faz diversas correções e lançamentos úteis durante a apuração do ICMS através da EFD ICMS/IPI.
O script percorre a EFD em busca de erros comuns e, quando encontrados, cria um popup perguntando se o usuário deseja que o script faça as correções.

As funções são:
 - Zerar valores de IPI.
 - Zerar valores de Abatimento Não Tributado.
 - Corrigir valores de Base de Cálculo maiores que o Valor de Operação (VL_BC <= VL_OP).
 - Corrigir valores de Redução da Base de Cálculo inconsistentes (VL_RED_BC = VL_OP - VL_BC).
 - Corrigir CSTs de importação usados equivocadamente (Corrige o primeiro dígito do CST de 1 para 2 e de 6 para 7).
 - Remove itens do registro 0200 não referenciados em nenhum registro posterior.
 - Remove duplicatas no registro de itens (0200) e no inventário (H010).
 - Corrige o valor do inventário (útil para ajustar inventários com valores minimamente divergentes).
 - Crédito do Simples Nacional: percorre os XMLs fornecidos e, se encontrados créditos do Simples Nacional, gera uma planilha de apuração e faz os lançamentos nos registros C197, aproveitando o crédito.
 - Bonificações: Sugere correções comparando os XMLs com as notas de entrada, indicando quando algum CFOP pode estar inconsistente.
 - Contadores: atualiza e corrige os registros contadores no final do arquivo de escrituração.

```pyinstaller .\SpedFix.py --onefile --noupx --noconsole -i .\icon.ico```

# Dependências
 - PySimpleGUI
 - XLSXWriter

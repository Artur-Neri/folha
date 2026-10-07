# Folha

Um leitor e editor leve de `.xlsx` para edições rápidas no navegador. Sem conta, sem assinatura e sem envio das planilhas para servidores.

**[Abrir o Folha no navegador](https://artur-neri.github.io/folha/)** · **[Baixar a versão offline](https://github.com/Artur-Neri/folha/releases/latest/download/index.html)**

## Abrir no Mac — sem instalar nada

1. [Baixe o HTML portátil da última release](https://github.com/Artur-Neri/folha/releases/latest/download/index.html). Se estiver compilando o projeto, use `portable/index.html`.
2. Renomeie para `Folha.html`, se quiser.
3. Abra o arquivo no Safari, Chrome, Edge ou Firefox. Você também pode arrastá-lo para uma janela do navegador.
4. Abra sua planilha, edite e clique em **Baixar .xlsx**.

O HTML contém o app e sua biblioteca de XLSX. Funciona offline, inclusive para abrir e exportar planilhas. Não precisa de Node no Mac. A versão portátil pesa cerca de 1 MB sem compressão.

## Funcionalidades

- Abrir `.xlsx` pelo seletor de arquivos ou arrastando na tela inicial.
- Criar uma planilha em branco e explorar um exemplo.
- Navegar entre abas e adicionar abas. Renomear com duplo clique em arquivos sem fórmulas.
- Editar com duplo clique, F2, digitando ou pela barra de fórmula.
- Selecionar arrastando ou com Shift + setas.
- Copiar e colar blocos de células (TSV), inclusive campos entre aspas e quebras de linha.
- Desfazer/refazer as últimas 100 operações de células.
- Negrito, itálico, formato numérico, moeda e porcentagem.
- Buscar conteúdo na aba e ir para qualquer célula pelo endereço.
- Soma e média de seleções numéricas.
- Download de uma nova cópia XLSX, mantendo o original intacto.
- Grade virtualizada nas duas direções: apenas as células visíveis entram no DOM.

## Atalhos

| Atalho | Ação |
| --- | --- |
| Enter / Tab | Confirmar edição e avançar |
| Escape | Cancelar edição |
| F2 / duplo clique | Editar célula |
| Setas / Shift + setas | Navegar / estender seleção |
| ⌘ ou Ctrl + C / V | Copiar / colar tabela |
| ⌘ ou Ctrl + Z | Desfazer |
| ⌘ ou Ctrl + Shift + Z | Refazer |
| ⌘ ou Ctrl + B | Negrito na grade |
| ⌘ ou Ctrl + F | Buscar na aba |
| ⌘ ou Ctrl + S | Baixar cópia |
| ⌘ ou Ctrl + O | Abrir arquivo |
| Delete / Backspace | Limpar conteúdo da seleção |

Use ponto ou vírgula para casas decimais, sem separador de milhar. Identificadores com zeros à esquerda continuam como texto. Comece com apóstrofo para forçar texto literal. Fórmulas começam com `=`. Datas existentes podem ser editadas no formato `AAAA-MM-DD`.

## Compatibilidade e limites

Usa ExcelJS para ler e escrever XLSX. Valores, datas, várias abas, fórmulas, mesclagens e estilos básicos são preservados em casos comuns. A visualização usa colunas de largura uniforme e não reproduz o layout de impressão ou as mesclagens visualmente. Formatos numéricos complexos são apresentados de maneira simplificada.

**Não é um substituto completo do Excel.** Não há motor de cálculo: fórmulas existentes exibem o resultado armazenado no arquivo, que pode ficar desatualizado após edições. Fórmulas novas só terão resultado depois de abrir em um aplicativo com cálculo. O download pede recálculo ao abrir no Excel.

Gráficos, imagens, macros, tabelas dinâmicas, vínculos externos e outros recursos avançados não têm garantia de preservação. Arquivos `.xls`, `.xlsm` e arquivos criptografados não são suportados. Para arquivos complexos, teste sempre com uma cópia. Não é permitido renomear abas de arquivos com fórmulas para evitar referências quebradas.

Adicionar/renomear uma aba reinicia o histórico de células; alterações de abas não são desfeitas. Operações de edição/cópia/colagem são limitadas a 50 mil células por vez. Arquivos são carregados inteiros na memória, portanto o tamanho viável depende do navegador e da máquina. A virtualização melhora a navegação, mas não reduz a memória necessária para ler o XLSX.

**Não há salvamento automático ou persistência após fechar a página.** Baixe suas alterações antes de sair. O app avisa ao tentar fechar com alterações pendentes; a confirmação de download significa que o navegador iniciou a transferência.

## Desenvolver / hospedar

Requer Node.js 22.12+ (ou 20.19+) somente para desenvolvimento/build.

```sh
npm ci
npm run dev
```

Abra o endereço exibido no terminal. Para gerar as versões web e portátil:

```sh
npm run build
```

- `dist/`: arquivos estáticos para hospedar. O parser de XLSX é carregado sob demanda.
- `portable/index.html`: versão autocontida, para abrir diretamente no navegador.

O site é publicado automaticamente no GitHub Pages a cada push na branch `main`, depois do build e dos testes. A versão web usa a base `/folha/`; ajuste `base` em `vite.config.ts` para hospedar em outro caminho.

Para verificar:

```sh
npx playwright install chromium
npm test
```

Os testes exercitam importação/exportação real com 20 mil linhas, edição, desfazer/refazer, preservação de tipos/estilos/fórmulas/mesclagens/abas, colagem, busca, seleção, salvamento pela barra de fórmula, telas pequenas e a versão portátil offline via `file://`.

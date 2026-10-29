# listar-arquivos

Script para Windows que lista todos os arquivos e pastas de um ou mais diretórios, incluindo as subpastas, e grava o resultado numa planilha CSV. Ele lê apenas os metadados (nome, tamanho, datas e atributos), sem abrir nenhum arquivo, então os arquivos **"somente online" do OneDrive não são baixados**.

A planilha serve para conferir o conteúdo de uma pasta ou comparar duas pastas, por exemplo um backup e a pasta atual.

## Requisitos

- Windows 10 ou 11
- Windows PowerShell 5.1, que já vem instalado no Windows

## Como usar

Há três formas de rodar o script:

1. **Dar dois cliques em `listar_arquivos.bat`.** Abre a janela "Selecionar pasta" do Windows, onde é possível escolher várias pastas de uma vez com Ctrl+clique ou Shift+clique. Depois o script pergunta se você quer adicionar pastas de outro local.
2. **Arrastar uma ou mais pastas do Explorer para cima de `listar_arquivos.bat`.**
3. **Pela linha de comando**, passando as pastas como argumentos:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\listar_arquivos.ps1 "C:\Users\fulano\OneDrive\Documentos" "E:\Backup\Documentos"
   ```

   Para gravar a planilha em outro lugar, use `-Destino`:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\listar_arquivos.ps1 "D:\Fotos" -Destino "D:\Inventarios"
   ```

Durante a execução, o script mostra o progresso e, para cada pasta, um resumo com:

- a quantidade de arquivos e o tamanho total;
- quantos arquivos são somente online;
- a quantidade de subpastas;
- o tempo gasto.

## Planilha gerada

Cada execução gera **um único CSV** com todas as pastas escolhidas, gravado em `inventarios/`:

- `inventario_<pasta>_<AAAAMMDD_HHMMSS>.csv` quando uma pasta é analisada (o nome da pasta vira minúsculas, sem acentos, com `_`);
- `inventario_<n>_pastas_<AAAAMMDD_HHMMSS>.csv` quando são várias.

O formato do arquivo:

- Separador `;` e codificação UTF-8. O arquivo abre direto no Excel ou no Bloco de Notas, com os acentos corretos.
- Uma planilha existente nunca é sobrescrita: se o nome já existir, o script acrescenta `_2`, `_3` e assim por diante.
- Os itens aparecem em ordem de árvore: cada pasta, depois os arquivos dela e depois as subpastas, em ordem alfabética.

| Coluna | Conteúdo |
|---|---|
| `PastaOrigem` | Pasta escolhida de onde o item veio (útil quando várias pastas são listadas juntas) |
| `Tipo` | `Arquivo`, `Pasta` ou `ERRO` (pasta que não pôde ser lida) |
| `CaminhoRelativo` | Caminho do item a partir da pasta de origem |
| `Nome` | Nome do arquivo ou da pasta |
| `PastaRelativa` | Subpasta onde o item está (vazio quando está direto na pasta de origem) |
| `Extensao` | Extensão do arquivo, como `.jpg` |
| `Nivel` | Profundidade: `1` para itens direto na pasta de origem, `2` para os de uma subpasta, e assim por diante |
| `TamanhoBytes` | Tamanho exato em bytes (só arquivos) |
| `Tamanho` | Tamanho legível, como `1,50 MB` |
| `Criado` / `Modificado` | Datas no formato `AAAA-MM-DD HH:MM:SS` |
| `Atributos` | Atributos do Windows, como `Archive, ReparsePoint` |
| `StatusOneDrive` | `Somente online`, `Sempre manter neste dispositivo` ou `Disponivel localmente` |
| `Observacao` | Mensagem de erro nas linhas `ERRO` |

### Detalhes

- **OneDrive:** os arquivos que estão só na nuvem continuam só na nuvem. A coluna `StatusOneDrive` mostra a situação de cada um.
- **Pastas que não puderam ser lidas,** por exemplo por falta de permissão, aparecem com `Tipo = ERRO`. O script avisa no final que o inventário está incompleto, para que isso não passe despercebido.
- **Caminhos com mais de 260 caracteres** são lidos normalmente.
- **Itens que não entram na planilha:**
  - a própria pasta escolhida, só o conteúdo dela é listado;
  - a planilha que está sendo gravada;
  - uma pasta escolhida duas vezes, que é listada uma vez só.

## Estrutura do repositório

```
listar_arquivos.bat   atalho: roda o script com dois cliques ou arrastando pastas
listar_arquivos.ps1   o script
inventarios/          planilhas geradas pelo script
relatorios/           relatórios de análises feitas a partir das planilhas
```

O conteúdo de `inventarios/` e `relatorios/` fica **fora do git** (veja o `.gitignore`), porque as planilhas e os relatórios trazem nomes de arquivos pessoais. As pastas só existem no repositório por causa do arquivo `.gitkeep`.

## Licença

Domínio público ([The Unlicense](LICENSE)).

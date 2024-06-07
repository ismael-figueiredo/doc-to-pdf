# Conversor de DOC para PDF

Este script Python automatiza a conversão de documentos do Microsoft Word (.doc) para o formato PDF. Utiliza a biblioteca `pywin32` para interagir com o Microsoft Word, garantindo que os documentos sejam abertos e convertidos de maneira adequada. É ideal para usuários que precisam converter vários documentos de uma vez, economizando tempo e esforço manual.

## Pré-requisitos

Para usar este script, você precisa ter:

- Python instalado em sua máquina (o script foi testado com Python 3.8).
- Microsoft Word instalado no seu sistema Windows.

## Instalação

Antes de executar o script, instale as bibliotecas necessárias. Abra um terminal e execute o seguinte comando para instalar as dependências do `requirements.txt`:

```bash
pip install -r requirements.txt
```
### Configuração

- **Diretório dos Arquivos:** Edite o script para apontar para o diretório onde seus arquivos .doc estão localizados. Altere a variável diretorio_raiz para o caminho do diretório correspondente.
####
- **Diretório de Saída:** Os arquivos PDF convertidos serão salvos em uma subpasta chamada pdf dentro do diretório especificado. O script criará esta pasta automaticamente se ela não existir.

### Uso
Para executar o script, navegue até o diretório onde o script está localizado e execute-o com Python através do terminal:

```bash
python index.py
```


### Funcionamento do Script
O script itera sobre todos os arquivos com a extensão .doc no diretório especificado. Para cada arquivo, ele:

• Abre o documento usando o Microsoft Word.
• Converte o documento para PDF.
• Salva o PDF na subpasta pdf com o mesmo nome do arquivo original.
• Mensagens de status serão impressas no terminal para cada arquivo convertido ou em caso de erros.

### Solução de Problemas
Se você encontrar problemas durante a execução do script, verifique:

• Se o Microsoft Word está corretamente instalado em sua máquina.
• Se os caminhos para os arquivos estão corretos e acessíveis pelo script.
• Se você tem permissões suficientes para ler os arquivos .doc e escrever nos diretórios.
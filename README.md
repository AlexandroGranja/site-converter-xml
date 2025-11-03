# Site Converter XML

Sistema web para processamento e seleção de arquivos XML de notas fiscais (NF-e) e conhecimentos de transporte (CT-e) com base em números de nota fiscal especificados em planilha Excel.

## 📋 Descrição

O **Site Converter XML** é uma aplicação web desenvolvida em Flask que permite processar arquivos XML de forma automatizada. O sistema compara números de notas fiscais de uma planilha Excel com arquivos XML compactados em ZIP e gera um novo arquivo ZIP contendo apenas os XMLs correspondentes.

### Funcionalidades Principais

- ✅ Upload de planilha Excel (.xlsx) com números de NF na coluna B
- ✅ Upload de arquivo ZIP contendo múltiplos XMLs
- ✅ Processamento automático de XMLs (NF-e e CT-e)
- ✅ Extração e correspondência de números de NF/CT
- ✅ Geração de arquivo ZIP com XMLs selecionados
- ✅ Interface web intuitiva e responsiva
- ✅ Processamento em memória (não salva arquivos temporários no disco)

## 🚀 Requisitos

- Python 3.8 ou superior
- Windows 10/11 (para usar o script `iniciar_windows.bat`)
- Navegador web moderno

## 📦 Instalação

### Método 1: Execução Automática (Windows)

1. Clone ou baixe este repositório
2. Execute o arquivo `iniciar_windows.bat`
3. O script irá:
   - Verificar se o Python está instalado
   - Criar um ambiente virtual automaticamente (se necessário)
   - Instalar as dependências necessárias
   - Iniciar o servidor web

### Método 2: Instalação Manual

1. Clone o repositório:
```bash
git clone https://github.com/AlexandroGranja/site-converter-xml.git
cd site-converter-xml
```

2. Crie e ative um ambiente virtual:
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

4. Execute a aplicação:
```bash
python -m src.main
```

## 🎯 Como Usar

1. **Inicie o servidor** (usando um dos métodos acima)
2. **Acesse a aplicação** no navegador em `http://localhost:5000`
3. **Prepare os arquivos**:
   - Planilha Excel (.xlsx) com os números de NF na **coluna B** (a partir da linha 2)
   - Arquivo ZIP contendo os XMLs de notas fiscais e conhecimentos de transporte
4. **Faça o upload**:
   - Selecione a planilha Excel
   - Selecione o arquivo ZIP com os XMLs
   - Clique em "Processar Arquivos"
5. **Aguarde o processamento** e baixe o arquivo ZIP resultante

### Formato da Planilha Excel

A planilha deve conter os números de NF na **coluna B**, começando da **linha 2** (a linha 1 pode ser cabeçalho).

Exemplo:
```
| Coluna A | Coluna B |
|----------|----------|
| Cabeçalho| NF       |
| Item 1   | 12345    |
| Item 2   | 67890    |
```

### Formato dos XMLs

O sistema suporta:
- **NF-e** (Nota Fiscal Eletrônica): arquivos XML com extensão `-nfe.xml`
- **CT-e** (Conhecimento de Transporte Eletrônico): arquivos XML com extensão `-cte.xml`

Os XMLs podem estar em qualquer estrutura de pastas dentro do arquivo ZIP.

## 🔧 Tecnologias Utilizadas

- **Python 3.8+**
- **Flask** - Framework web
- **openpyxl** - Leitura de arquivos Excel
- **xml.etree.ElementTree** - Processamento de XML
- **zipfile** - Manipulação de arquivos ZIP
- **Bootstrap 5** - Interface frontend

## 📁 Estrutura do Projeto

```
site-converter-xml/
├── src/
│   ├── __init__.py
│   ├── main.py                 # Aplicação principal Flask
│   ├── automation_tasks.py      # Automação para download do Painel Target
│   └── templates/
│       └── index.html          # Interface web
├── requirements.txt            # Dependências Python
├── iniciar_windows.bat         # Script de inicialização (Windows)
├── .env                        # Variáveis de ambiente (se necessário)
└── README.md                   # Este arquivo
```

## 🔍 Como Funciona

1. **Leitura da Planilha**: O sistema lê a coluna B do arquivo Excel e extrai os números de NF
2. **Extração do ZIP**: Descompacta o arquivo ZIP na memória
3. **Análise dos XMLs**: Para cada XML, extrai o número da NF ou CT-e
4. **Correspondência**: Verifica se o número da NF/CT termina com algum dos números da planilha
5. **Seleção**: Copia apenas os XMLs correspondentes
6. **Geração do Resultado**: Cria um novo arquivo ZIP com os XMLs selecionados

### Algoritmo de Correspondência

O sistema verifica se o número completo da NF/CT-e **termina** com os números especificados na planilha. Isso permite encontrar correspondências mesmo quando há prefixos diferentes.

Exemplo:
- Planilha: `12345`
- XML: `000012345` → ✅ **Corresponde**
- XML: `9999912345` → ✅ **Corresponde**
- XML: `123456` → ❌ **Não corresponde**

## 🛠️ Configurações

### Limites de Upload

- Tamanho máximo por arquivo: **50 MB**
- Tipos permitidos: `.xlsx` e `.zip`

### Porta do Servidor

Por padrão, o servidor roda na porta **5000**. Para alterar, edite o arquivo `src/main.py`:

```python
app.run(debug=True, host='0.0.0.0', port=5000)  # Altere a porta aqui
```

## 🐛 Solução de Problemas

### Erro: "Python não encontrado"
- Instale o Python 3.8 ou superior
- Certifique-se de que o Python está no PATH do sistema
- Baixe em: https://www.python.org/downloads/

### Erro: "Nenhum número de NF encontrado na coluna B"
- Verifique se os números estão na coluna B (não na coluna A)
- Certifique-se de que há dados a partir da linha 2
- Verifique se o arquivo é um .xlsx válido

### Erro: "Nenhum XML correspondente encontrado"
- Verifique se os números da planilha estão corretos
- Confirme que os XMLs estão dentro do arquivo ZIP
- Verifique se os XMLs são de NF-e ou CT-e válidos

### Servidor não inicia
- Verifique se a porta 5000 está disponível
- Feche outras aplicações que possam estar usando a porta
- Execute como administrador se necessário

## 📝 Notas Importantes

- ⚠️ O sistema processa arquivos **em memória**, não salva temporários no disco
- ⚠️ Arquivos grandes (>50MB) podem causar lentidão
- ⚠️ O servidor roda em modo debug por padrão (não use em produção)

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📄 Licença

Este projeto é privado e destinado ao uso interno.

## 👤 Autor

**Alexandro Granja**

- GitHub: [@AlexandroGranja](https://github.com/AlexandroGranja)

---

Desenvolvido com ❤️ usando Flask e Python


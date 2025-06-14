# 📝 WordFixer.AI

WordFixer.AI é uma ferramenta desenvolvida em C# para leitura, extração, correção e substituição de textos em arquivos Word (.docx), preservando toda a formatação do documento.

Utiliza a API da OpenAI para corrigir erros gramaticais, ortográficos e de concordância textual de forma automatizada. Ideal para programadores, revisores ou equipes que precisam corrigir documentos em larga escala sem alterar o layout.

---

## ✨ Funcionalidades

- **Leitura de documentos .docx**: Utiliza OpenXML para manipulação de arquivos Word.
- **Extração de texto**: Estrutura os textos por ID para facilitar o processamento.
- **Correção automática**: Usa IA (OpenAI GPT) para corrigir erros gramaticais e de concordância.
- **Substituição segura**: Insere os textos corrigidos no documento original sem alterar a formatação.
- **Preservação de estilos**: Mantém parágrafos, títulos, estilos e outras formatações intactas.

---

## 🔧 Tecnologias

- **C#** (.NET 8 Console App)
- **DocumentFormat.OpenXml**: Para manipulação de documentos Word.
- **OpenAI API**: Para correção de textos com inteligência artificial.
- **Newtonsoft.Json**: Para manipulação de JSON.

---

## 🚀 Como usar

1. **Preparação**:
   - Certifique-se de que o arquivo `config.json` contém sua chave da API OpenAI.
   - Coloque o arquivo `.docx` na pasta do projeto.

2. **Execução**:
   - Execute o programa para extrair os textos do documento Word.
   - Envie os textos para a API OpenAI para correção.
   - Aplique as correções no documento original, preservando a formatação.

3. **Resultado**:
   - O documento corrigido será salvo com todas as alterações aplicadas.

---

## 📂 Estrutura do Projeto

- `Program.cs`: Contém a lógica principal para leitura e correção de textos.
- `Service/OpenAIService.cs`: Implementa a integração com a API OpenAI.
- `config.json`: Arquivo de configuração contendo a chave da API.
- `.gitignore`: Arquivo para ignorar arquivos desnecessários no repositório.

---

## 🛠️ Configuração

Certifique-se de que o arquivo `config.json` está configurado corretamente:


---

## 📝 Exemplo de Uso

### Entrada:
Arquivo Word com texto:


---

## 📝 Exemplo de Uso

### Entrada:
Arquivo Word com texto:



---

## 📄 Licença

Este projeto está licenciado sob a [MIT License](https://opensource.org/licenses/MIT).

---

## 🤝 Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.

---

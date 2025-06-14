using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Threading.Tasks;
using WordFixer.Console.Service;
using System.IO;
using DocumentFormat.OpenXml;

internal class Program
{
    static async Task Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8; // Configura o console para UTF-8
        Console.WriteLine("🔄 Iniciando o processo de correção de texto...");

        var diretorio = "D:\\Relatorios\\";
        Console.WriteLine($"📂 Diretório definido: {diretorio}");

        Console.WriteLine("📄 Abrindo o arquivo Word para correção...");
        var resultado = await CorrigirBlocosTexto($"{diretorio}1.docx");

        Console.WriteLine("💾 Salvando logs no arquivo...");
        string logFilePath = $"{diretorio}resultado_logs.txt"; // Caminho do arquivo de log

        using (StreamWriter writer = new StreamWriter(logFilePath, append: true))
        {
            foreach (var log in resultado.Logs)
            {
                writer.WriteLine(log); // Escreve cada log no arquivo
                Console.WriteLine(log); // Mantém a exibição no console
            }
        }

        Console.WriteLine($"\n✅ Logs salvos em: {Path.GetFullPath(logFilePath)}");

        if (resultado.Erros.Count > 0)
        {
            Console.WriteLine("\n❌ Erros encontrados durante o processo:");
            foreach (var e in resultado.Erros)
                Console.WriteLine(e);
        }
        else
        {
            Console.WriteLine("\n✅ Todas as tags foram balanceadas com sucesso.");
        }

        Console.WriteLine("🏁 Processo concluído.");
    }

    public class Resultado
    {
        public List<string> Logs { get; set; } = new();
        public List<string> Erros { get; set; } = new();
    }

    public static async Task<Resultado> CorrigirBlocosTexto(string caminho)
    {
        Console.WriteLine($"📄 Abrindo o documento: {caminho}");
        var resultado = new Resultado();
        if (!File.Exists(caminho))
        {
            Console.WriteLine($"❌ O arquivo especificado não foi encontrado: {caminho}");
            resultado.Erros.Add($"❌ O arquivo especificado não foi encontrado: {caminho}");
            return resultado; // Retorna imediatamente para evitar exceções
        }
        try
        {
            // Tenta abrir o arquivo para verificar se está acessível
            using (FileStream fs = new FileStream(caminho, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Console.WriteLine("✅ O arquivo está acessível.");
            }
        }
        catch (IOException)
        {
            Console.WriteLine($"❌ O arquivo está sendo usado por outro processo: {caminho}");
            resultado.Erros.Add($"❌ O arquivo está sendo usado por outro processo: {caminho}");
            return resultado; // Retorna imediatamente para evitar exceções
        }

        Console.WriteLine($"📄 Abrindo o documento: {caminho}");

        using var doc = WordprocessingDocument.Open(caminho, true);
        var body = doc.MainDocumentPart.Document.Body;

        Console.WriteLine("🔍 Extraindo todos os nodes de texto...");
        var allTextNodes = body.Descendants<Text>().ToList();

        bool dentroDoBloco = false;
        List<Text> buffer = new();
        List<(List<Text> texts, string textoOriginal)> blocos = new();

        Console.WriteLine("🔄 Processando os nodes de texto...");
        foreach (var textNode in allTextNodes)
        {
            var texto = textNode.Text;

            if (texto.Contains("[INICIO]"))
            {
                if (dentroDoBloco)
                {
                    resultado.Erros.Add("❌ Tag [INICIO] duplicada sem fechar.");
                }
                dentroDoBloco = true;
                buffer = new(); // iniciar novo bloco
                Console.WriteLine("🔧 Início de um novo bloco detectado.");
            }

            if (dentroDoBloco && texto != "[INICIO]" && texto != "[FIM]")
                buffer.Add(textNode);

            if (texto.Contains("[FIM]"))
            {
                if (!dentroDoBloco)
                {
                    resultado.Erros.Add("❌ Tag [FIM] sem [INICIO] correspondente.");
                }
                else
                {
                    dentroDoBloco = false;
                    blocos.Add((new List<Text>(buffer), string.Concat(buffer.Select(t => t.Text))));
                    Console.WriteLine("🔧 Fim de um bloco detectado.");
                }
            }
        }

        if (dentroDoBloco)
            resultado.Erros.Add("❌ TAG do Bloco iniciado com [INICIO] mais não foi encontrado a TAG [FIM].");

        if (blocos.Count == 0)
            resultado.Erros.Add("❌ Nenhum bloco foi informado com [INICIO] e [FIM].");

        Console.WriteLine($"🔄 Total de blocos detectados: {blocos.Count}");

        foreach (var (texts, original) in blocos)
        {
            Console.WriteLine("🔄 Corrigindo texto do bloco...");
            int correcao = 0;
            TextoExtraido textoExtraido;

            do
            {
                textoExtraido = await CorrigirTexto(original);
            } while (original.Length - 50 > textoExtraido.TextoNovo.Length);
            string corrigido = textoExtraido.TextoNovo;

            resultado.Logs.Add(
                $" 🔧 {correcao++}) Referencia: {corrigido.Substring(0, 35)}\nErros encontrados: {textoExtraido.PalavraComErro}");

            int pos = 0;
            for (int i = 0; i < texts.Count; i++)
            {
                var t = texts[i];

                if (pos >= corrigido.Length)
                {
                    t.Text = "";
                    continue;
                }

                int tam = t.Text.Length;

                if (tam > corrigido.Length - pos)
                    tam = corrigido.Length - pos;

                var condicao = tam + pos < corrigido.Length && corrigido[tam + pos] == ' ';
                // Inclui o espaço, se houver, no texto atual
                if (condicao)
                {
                    tam++;
                }

                // Enquanto o caractere atual não for um espaço e não ultrapassar o tamanho da string corrigida,
                // incrementa o tamanho (tam) para incluir mais caracteres.
                while (tam + pos < corrigido.Length && corrigido[tam + pos] != ' ')
                {
                    tam++;
                }
                var textoAtual = corrigido.Substring(pos, tam);
                t.Space = SpaceProcessingModeValues.Preserve;
                t.Text = textoAtual;
                pos += tam;
            }

            // Se o texto corrigido exceder o número de nodes, crie novos nodes
            if (pos < corrigido.Length)
            {
                var parent = texts.First().Parent;
                while (pos < corrigido.Length)
                {
                    var newTextNode = new Text(corrigido.Substring(pos, Math.Min(50, corrigido.Length - pos)))
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    };
                    parent.AppendChild(new Run(newTextNode));
                    pos += newTextNode.Text.Length;
                }
            }
        }

        Console.WriteLine("💾 Salvando alterações no documento...");
        doc.MainDocumentPart.Document.Save();
        Console.WriteLine("✅ Alterações salvas com sucesso.");

        return resultado;
    }

    static async Task<TextoExtraido> CorrigirTexto(string texto)
    {
        Console.WriteLine("🔄 Enviando texto para correção via API...");
        var configContent = File.ReadAllText("config.json");
        if (string.IsNullOrWhiteSpace(configContent))
        {
            throw new InvalidOperationException("❌ O arquivo de configuração está vazio ou ausente.");
        }

        var config = JsonConvert.DeserializeObject<Dictionary<string, string>>(configContent);
        if (config == null || !config.ContainsKey("OpenAIKey"))
        {
            throw new InvalidOperationException("❌ O arquivo de configuração não contém a chave 'OpenAIKey'.");
        }

        string apiKey = config["OpenAIKey"];

        var nvTexto = await new OpenAIService(apiKey)
            .CorrigirTextosAsync(texto);

        Console.WriteLine("✅ Texto corrigido com sucesso.");
        return nvTexto;
    }
}

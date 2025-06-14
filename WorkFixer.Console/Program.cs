//var listaCorrigida = CorrigirDocumentoWord("D:\\Relatorios\\RDIC 2° PERÍODO B - 1º semestre.docx");

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Threading.Tasks;
using WordFixer.Console.Service;
using System.IO;
internal class Program
{
    static async Task Main(string[] args)
    {


        var resultado = await CorrigirBlocosTexto("D:\\Relatorios\\RDIC 2° PERÍODO B - 1º semestre.docx");

        Console.WriteLine("\nResultado:");
        foreach (var log in resultado.Logs)
            Console.WriteLine(log);

        if (resultado.Erros.Count > 0)
        {
            Console.WriteLine("\n❌ Erros:");
            foreach (var e in resultado.Erros)
                Console.WriteLine(e);
        }
        else
        {
            Console.WriteLine("\n✅ Todas as tags balanceadas.");
        }
    }

    public class Resultado
    {
        public List<string> Logs { get; set; } = new();
        public List<string> Erros { get; set; } = new();
    }

    public static async Task<Resultado> CorrigirBlocosTexto(string caminho)
    {
        var resultado = new Resultado();

        using var doc = WordprocessingDocument.Open(caminho, true);
        var body = doc.MainDocumentPart.Document.Body;

        var allTextNodes = body.Descendants<Text>().ToList();

        bool dentroDoBloco = false;
        List<Text> buffer = new();
        List<(List<Text> texts, string textoOriginal)> blocos = new();

        foreach (var textNode in allTextNodes)
        {
            var texto = textNode.Text;

            if (texto.Contains("##INICIO##"))
            {
                if (dentroDoBloco)
                {
                    resultado.Erros.Add("❌ Tag ##INICIO## duplicada sem fechar.");
                }
                dentroDoBloco = true;
                buffer = new(); // iniciar novo bloco
            }

            if (dentroDoBloco && texto != "##INICIO##" && texto != "##FIM##")
                buffer.Add(textNode);

            if (texto.Contains("##FIM##"))
            {
                if (!dentroDoBloco)
                {
                    resultado.Erros.Add("❌ Tag ##FIM## sem ##INICIO## correspondente.");
                }
                else
                {
                    dentroDoBloco = false;
                    blocos.Add((new List<Text>(buffer), string.Concat(buffer.Select(t => t.Text))));
                }
            }
        }

        if (dentroDoBloco)
            resultado.Erros.Add("❌ Bloco iniciado com ##INICIO## sem ##FIM##.");

        if (blocos.Count == 0)
            resultado.Erros.Add("❌ Bloco não informado com ##INICIO## e ##FIM##.");

        foreach (var (texts, original) in blocos)
        {
            string corrigido = await CorrigirTexto(original);
            resultado.Logs.Add($"\n🔧 Bloco corrigido:\n↳ Original : {original}\n↳ Corrigido: {corrigido}");

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

                // Enquanto o caractere atual não for um espaço e não ultrapassar o tamanho da string corrigida,
                // incrementa o tamanho (tam) para incluir mais caracteres.
                while (tam + pos < corrigido.Length && corrigido[tam + pos] != ' ')
                {
                    tam++;
                }

                // Inclui o espaço, se houver, no texto atual
                if (tam + pos < corrigido.Length && corrigido[tam + pos] == ' ')
                {
                    tam++;
                }

                t.Text = corrigido.Substring(pos, tam);
                pos += tam;

            }
        }

        doc.MainDocumentPart.Document.Save();
        return resultado;
    }

    static async Task<string> CorrigirTexto(string texto)
    {
        var configContent = File.ReadAllText("config.json");
        if (string.IsNullOrWhiteSpace(configContent))
        {
            throw new InvalidOperationException("The configuration file is empty or missing.");
        }

        var config = JsonConvert.DeserializeObject<Dictionary<string, string>>(configContent);
        if (config == null || !config.ContainsKey("OpenAIKey"))
        {
            throw new InvalidOperationException("The configuration file does not contain the required 'OpenAIKey'.");
        }

        string apiKey = config["OpenAIKey"];

        var nvTexto = await new OpenAIService(apiKey)
            .CorrigirTextosAsync(texto);
        return nvTexto.TextoNovo;
    }
}

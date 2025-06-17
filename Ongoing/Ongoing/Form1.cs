using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using DocumentFormat.OpenXml.Drawing;





using System.IO;
using DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Presentation;
//using DocumentFormat.OpenXml.Presentation;
//using DocumentFormat.OpenXml.Drawing;


namespace Ongoing
{
    public partial class Form1 : Form
    {
        public Form1()
        {



            InitializeComponent();
        }
        static int posicao = 1;
        private void button1_Click(object sender, EventArgs e)
        {
            var caminhoModelo = @"C:\\Users\\80105812.REDECORP\\Desktop\\PowerPoint\\Entrada\\Ongoing Telco.pptx";
            var caminhoSaida = @"C:\Users\80105812.REDECORP\Desktop\PowerPoint\Saida\ApresentacaoGerada.pptx";
            var caminhoImagem = @"C:\Image\ficha_teste.jpg";
            var valoresLinha = new List<string> { "PP - PTI - 5797 - Novas Melhorias no Processo do Batimento – CEMI X SICS X EIR", "Estimativa", "02/02/2025", "TBD" };
            var valoresLinha2 = new List<string> { "PP - PSS - 484 - Ativar Flag de MFA do Siebel 15 dentro do ambiente do OAM", "Desenvolvimento", "02/01/2025", "TBD" };

            var valoresLinha3 = new List<string> { "Estruturante - teste - Novas Melhorias no Processo do teste", "Estimativa", "02/02/2025", "TBD" };
            var valoresLinha4 = new List<string> { "Estruturante - 484 - Ativar Flag de MFA do teste teste", "Desenvolvimento", "02/01/2025", "TBD" };
            //AdicionarLinhaTabela(caminhoSaida, 0, valoresLinha);
            //AdicionarLinhaTabela(caminhoSaida, 0, valoresLinha2);

            AdicionarLinha(caminhoModelo, caminhoSaida, valoresLinha, "ONGOING | Andamento PPs", 0);
            
            AdicionarLinha(caminhoModelo, caminhoSaida, valoresLinha2, "ONGOING | Andamento PPs", 0);
            var sss = ObterPosicaoDoSlide(caminhoSaida, "ONGOING | Andamento PPs");
            //AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoSaida, caminhoImagem, sss + 1);
          //  AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoSaida, caminhoImagem, sss + 2);

           // ImagemTitulo.CriarPptComTituloEImagem(caminhoSaida, caminhoImagem, sss + 1);
          


            // AdicionarImagem2.InserirSlideComImagem(caminhoSaida, caminhoImagem);

            AdicionarLinha(caminhoModelo, caminhoSaida, valoresLinha3, "ONGOING | Andamento Estruturante", 0);            
            AdicionarLinha(caminhoModelo, caminhoSaida, valoresLinha4, "ONGOING | Andamento Estruturante", 0);
            //AdicionarImagem2.InserirSlideComImagem(caminhoSaida, caminhoImagem);
            //AdicionarImagem2.InserirSlideComImagem(caminhoSaida, caminhoImagem);
            var sssss = ObterPosicaoDoSlide(caminhoSaida, "ONGOING | Andamento Estruturante");
           // AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoSaida, caminhoImagem, sssss + 1);
//AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoSaida, caminhoImagem, sssss + 2);


            MessageBox.Show("Hello, world.");
        }

        
        public static void AdicionarLinha(string caminhoModelo, string caminhoSaida, List<string> dadosTabela, string tabelaNome, int slideIndex)
        {
            var caminhoImagem = @"C:\Users\80105812.REDECORP\Desktop\PowerPoint\ficha_teste.jpg";
            //System.IO.File.Copy(caminhoModelo, caminhoSaida, true);
            int i = 0;
            bool inserirImagem = false;
            try
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoSaida, true))
                {
                    PresentationPart presentationPart = presentationDoc.PresentationPart;
                    var slides = presentationDoc.PresentationPart.SlideParts;
                    foreach (var slide in slides)
                    {
                        if (slide.Slide.InnerText.Contains(tabelaNome))
                        {
                             
                            // Obtém o slide desejado (começa em 0)
                            var slidePart = presentationDoc.PresentationPart.SlideParts.ElementAt(slideIndex);
                            // Busca a tabela no slide
                            var tabela = slide.Slide.Descendants<Table>().FirstOrDefault();
                            if (tabela == null)
                            {
                                Console.WriteLine("Tabela não encontrada no slide especificado.");
                                return;
                            }
                            // Obtém o número de colunas da tabela
                            int numColunas = tabela.TableGrid.ChildElements.Count;
                            // Cria uma nova linha
                            var novaLinha = new TableRow();
                            novaLinha.Height = 370000; // Altura da linha em EMUs
                                                       // Adiciona células à nova linha

                            foreach (var valor in dadosTabela)
                            {
                                var novaCelula = new TableCell();
                                // Adiciona o texto à célula
                                novaCelula.Append(new DocumentFormat.OpenXml.Drawing.TextBody(
                                    new BodyProperties(),

                                    new ListStyle(),
                                     //new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text(valor)))
                                     new Paragraph(
                                new Run(
                                    new RunProperties() { FontSize = 12 * 100 }, // Define o tamanho da fonte
                                    new DocumentFormat.OpenXml.Drawing.Text(valor)
                                )
                            )
                                ));
                                // Define a margem padrão
                                novaCelula.Append(new TableCellProperties());

                                novaLinha.Append(novaCelula);
                            }
                            // Garante que o número de células seja igual ao número de colunas
                            while (novaLinha.ChildElements.Count < numColunas)
                            {
                                novaLinha.Append(new TableCell(new TableCellProperties()));
                            }
                            // Adiciona a nova linha à tabela
                            tabela.Append(novaLinha);
                            inserirImagem = true;
                            break;
                        }
                        else
                        {
                            if(posicao<=1)
                                posicao++;
                        }
                            
                    }




                }
            }
            catch (Exception ex) { }
        }




        public static int ObterPosicaoDoSlide(string caminhoPptx, string tituloSlide)
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, false))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                if (presentationPart == null || presentationPart.Presentation == null)
                    throw new InvalidDataException("O arquivo PPTX não possui uma estrutura válida.");
                var slideIdList = presentationPart.Presentation.SlideIdList;
                int posicao = 1; // A posição começa em 1
                foreach (var slideId in slideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    // Verifica o título do slide (caso exista)
                    string titulo = ObterTituloDoSlide(slidePart);
                    if (!string.IsNullOrEmpty(titulo) && titulo == tituloSlide)
                    {
                        return posicao; // Retorna a posição do slide
                    }
                    posicao++;
                }
            }
            return -1; // Retorna -1 se o slide não for encontrado
        }

        private static string ObterTituloDoSlide(SlidePart slidePart)
        {
            if (slidePart == null || slidePart.Slide == null)
                return null;
            var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().FirstOrDefault(s => s.TextBody != null);
            if (shape.TextBody.InnerText != null)
            {
                return shape.TextBody.InnerText.Trim();
            }
            return null;
        }




    }
}

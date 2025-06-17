using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ongoing
{
    public class AdicionarImagem
    {
    

        public static void InserirSlideComImagemNaOrdem(string caminhoPptx, string caminhoImagem, string titulo, int posicaoSlide)
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, true))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                if (presentationPart == null || presentationPart.Presentation == null)
                    throw new InvalidDataException("O arquivo PPTX não possui uma estrutura válida.");
                // Adiciona um novo SlidePart
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                // Adiciona imagem ao slide
                AdicionarImagemAoSlide(slidePart, caminhoImagem, titulo);
                // Cria um novo SlideId
                uint novoSlideId = ObterProximoSlideId(presentationPart.Presentation);
                string relId = presentationPart.GetIdOfPart(slidePart);
                SlideId novoSlideIdElement = new SlideId() { Id = novoSlideId, RelationshipId = relId };
                // Adiciona o novo SlideId na posição especificada
                var slideIdList = presentationPart.Presentation.SlideIdList;
                if (posicaoSlide <= 0 || posicaoSlide > slideIdList.Count())
                {
                    // Adiciona ao final se a posição for inválida
                    slideIdList.Append(novoSlideIdElement);
                }
                else
                {
                    // Insere na posição especificada
                    var referenciaSlide = slideIdList.Elements<SlideId>().ElementAt(posicaoSlide - 1);
                    slideIdList.InsertBefore(novoSlideIdElement, referenciaSlide);
                }
                presentationPart.Presentation.Save();
            }
        }

        private static void AdicionarImagemAoSlide(SlidePart slidePart, string caminhoImagem, string NomeProjeto)
        {
            // Dimensões do slide
            long larguraSlide = 12111100;
            long alturaSlide = 6858000;

            // Altura ocupada pelo título
            long alturaTitulo = 571600;

            // Altura restante para a imagem
            long alturaImagem = alturaSlide - alturaTitulo;

            // Adiciona a imagem ao slide
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(caminhoImagem, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            // Define as dimensões da imagem
            long x = 1 * 914400; // 1 polegada da esquerda
            long y = 1 * 914400; // 1 polegada do topo
            long cx = 5 * 914400; // 5 polegadas de largura
            long cy = 6 * 914400; // 4 polegadas de altura
                                  // Adiciona a imagem ao slide
            string imagePartId = slidePart.GetIdOfPart(imagePart);
            var picture = new Picture(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties() { Id = 1U, Name = "Imagem" },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip() { Embed = imagePartId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 0, Y = alturaTitulo },
                    new A.Extents() { Cx = larguraSlide, Cy = alturaImagem }
                        //new A.Offset() { X = x, Y = y },
                        //new A.Extents() { Cx = cx, Cy = cy }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            Shape titulo = new Shape(
           new NonVisualShapeProperties(
               new NonVisualDrawingProperties() { Id = 1U, Name = "Título" },
               new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
               new ApplicationNonVisualDrawingProperties()),
           new ShapeProperties(
               new A.Transform2D(
                   new A.Offset() { X = 0, Y = 0 }, // posição
                   new A.Extents() { Cx = larguraSlide, Cy = alturaTitulo } // tamanho
                 //new A.Extents() { Cx = 9999990, Cy = 571600 } // tamanho
                   //new A.Extents() { Cx = 9144000, Cy = 1371600 }
                                                                 //new A.Offset() { X = 914400, Y = 400000 }, // posição
                                                                 //new A.Extents() { Cx = 7315200, Cy = 1000000 } // tamanho
               ),
               new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
           ),
           new TextBody(
               new A.BodyProperties(),
               new A.ListStyle(),
               new A.Paragraph(
                   new A.Run(
                       new A.RunProperties() { FontSize = 1600, Bold = false },
                       new A.Text(NomeProjeto)
                   )
               )
           )
       );
            // Adiciona a imagem ao ShapeTree do slide
            slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(picture);
            slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(titulo);
            slidePart.Slide.Save();
        }
        private static uint ObterProximoSlideId(Presentation presentation)
        {
            uint maxId = 255;
            if (presentation.SlideIdList != null)
            {
                foreach (SlideId slideId in presentation.SlideIdList.Elements<SlideId>())
                {
                    if (slideId.Id > maxId)
                        maxId = slideId.Id;
                }
            }
            return maxId + 1;
        }
    }
}

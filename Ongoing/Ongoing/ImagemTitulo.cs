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
using P = DocumentFormat.OpenXml.Presentation;

namespace Ongoing
{
    public class ImagemTitulo
    {
        public static void CriarPptComTituloEImagem(string caminhoPptx, string caminhoImagem, int posicaoSlide)
        {
             

            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, true))            
            {
                //PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                //presentationPart.Presentation = new Presentation();

                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

                //SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                //slideMasterPart.SlideMaster = new SlideMaster();
                //SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                //slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));
                //slideMasterPart.SlideMaster.Append(new SlideLayoutIdList(new SlideLayoutId() { Id = 1, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }));
                //presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(new SlideMasterId() { Id = 1, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });

                //SlideIdList slideIdList = new SlideIdList();
                //uint slideId = 256U;
                //slideIdList.Append(new SlideId() { Id = slideId, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
                //presentationPart.Presentation.Append(slideIdList);
                //presentationPart.Presentation.Save();

                // Adiciona o título como TextBox
                AddTituloNoSlide(slidePart, "titulossssssssss");

                // Adiciona a imagem
                ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream stream = new FileStream(caminhoImagem, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                string rId = slidePart.GetIdOfPart(imagePart);
                AddImagemNoSlide(slidePart, rId);

                slidePart.Slide.Save();
            }

            Console.WriteLine("Slide criado com título e imagem: " + caminhoPptx);
        }

        private static void AddTituloNoSlide(SlidePart slidePart, string titulo)
        {
            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            Shape titleShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = shapeId, Name = "Title" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties() { FontSize = 4400, Bold = true }, // 22pt
                            new A.Text(titulo)
                        ),
                        new A.EndParagraphRunProperties() { Language = "pt-BR", Dirty = false }
                    )
                )
            );

            // Define a posição e tamanho do título
            titleShape.ShapeProperties = new ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = 914400, Y = 400000 }, // Centralizado no topo
                    new A.Extents() { Cx = 7315200, Cy = 1000000 } // Largura e altura do título
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );

            shapeTree.AppendChild(titleShape);
        }

        private static void AddImagemNoSlide(SlidePart slidePart, string relationshipId)
        {
            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            uint picId = 2;

            Picture picture = new Picture(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties() { Id = picId, Name = "Imagem" },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new BlipFill(
                    new A.Blip() { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 914400, Y = 1500000 }, // Posição da imagem
                        new A.Extents() { Cx = 5000000, Cy = 4000000 } // Tamanho da imagem
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            shapeTree.AppendChild(picture);
        }
    }
}

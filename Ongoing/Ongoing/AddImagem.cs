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
    public class AddImagem
    {

        public static void InserirNovoSlideComImagem(string caminhoPptx, string caminhoImagem)
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, true))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                if (presentationPart == null || presentationPart.Presentation == null)
                    return;
                // Cria novo slide
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                // Adiciona imagem
                ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream stream = new FileStream(caminhoImagem, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                string imagePartId = slidePart.GetIdOfPart(imagePart);
                long x = 1 * 914400;
                long y = 1 * 914400;
                long cx = 4 * 914400;
                long cy = 3 * 914400;
                var picture = new Picture(
                    new NonVisualPictureProperties(
                        new NonVisualDrawingProperties() { Id = 1U, Name = "Nova Imagem" },
                        new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new BlipFill(
                        new A.Blip() { Embed = imagePartId },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new ShapeProperties(
                        new A.Transform2D(
                            new A.Offset() { X = x, Y = y },
                            new A.Extents() { Cx = cx, Cy = cy }
                        ),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                    )
                );
                slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(picture);
                slidePart.Slide.Save();
                // Adiciona o novo slide à apresentação
                uint newSlideId = ObterProximoSlideId(presentationPart.Presentation);
                string relId = presentationPart.GetIdOfPart(slidePart);
                SlideId slideId = new SlideId() { Id = newSlideId, RelationshipId = relId };
                if (presentationPart.Presentation.SlideIdList == null)
                {
                    presentationPart.Presentation.SlideIdList = new SlideIdList();
                }
                presentationPart.Presentation.SlideIdList.Append(slideId);
                presentationPart.Presentation.Save();
            }
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

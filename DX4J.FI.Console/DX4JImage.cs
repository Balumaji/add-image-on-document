using java.io;
using org.docx4j.dml.wordprocessingDrawing;
using org.docx4j.jaxb;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts.relationships;
using org.docx4j.openpackaging.parts.WordprocessingML;
using org.docx4j.wml;
using System;

namespace DX4J.Image
{
    /// <summary>
    /// <author>Balaji</author>
    /// </summary>
    public class DX4JImage
    {
        public enum ImageLayoutOptions
        {
            INLINE_WITH_TEXT,
            ANCHOR_RELATIVE_POS
        }

        private WordprocessingMLPackage wordMLPackage;
        private string inputFile;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fileWordDocument">Word document to be added with the image with full path</param>
        public DX4JImage(string fileWordDocument)
        {
            PrepareRuntime();
            this.inputFile = fileWordDocument;
            try
            {
                this.wordMLPackage = WordprocessingMLPackage.load(new java.io.File(fileWordDocument));
            }
            catch(Exception ex)
            {
                throw new ApplicationException("Fatal error on opening the word document!!");
            }        
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="imageFile">Image file with path</param>
        /// <param name="textOnDocument">Text which is unique on where the image will be added</param>
        public void AddImageInline(
            string imageFile,
            string textOnDocument
            )
        {
            P p = SearchParaWithText(this.wordMLPackage, textOnDocument);
            if (p != null)
                CreateInlineImageObject(wordMLPackage, imageFile, ref p);
            else
                throw new ApplicationException("Error on locating the search text!!");

            wordMLPackage.save(new File(this.inputFile));
        }

        /// <summary>
        /// To load IKVM assemblies during runtime
        /// </summary>
        private void PrepareRuntime()
        {
            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(typeof(org.docx4j.jaxb.Context)));
            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(typeof(org.apache.xalan.processor.TransformerFactoryImpl)));
        }

        /// <summary>
        /// Add image as Anchor, with respect to positions on the page with position defined in inches.
        /// The 'textOnDocument' will be the search key to identify the page on which the image will be added
        /// </summary>
        /// <param name="imageFile"></param>
        /// <param name="textOnDocument"></param>
        /// <param name="verticalPos"></param>
        /// <param name="horizontalPos"></param>
        public void AddImageAnchor(
            string imageFile,
            string textOnDocument,
            int verticalPos,
            int horizontalPos
            )
        {

            P p = SearchParaWithText(this.wordMLPackage, textOnDocument);
            if (p != null)
                CreateAnchorImageObject(wordMLPackage, imageFile, verticalPos, horizontalPos, ref p);
            else
                throw new ApplicationException("Error on locating the search text!!");

            wordMLPackage.save(new File(this.inputFile));
        }

        /// <summary>
        /// To identify the para, which holds the search text.
        /// </summary>
        /// <param name="wordMLPackage"></param>
        /// <param name="searchText"></param>
        /// <returns></returns>
        private P SearchParaWithText(WordprocessingMLPackage wordMLPackage, string searchText)
        {
            P p = null;

            org.docx4j.wml.Document doc = (org.docx4j.wml.Document)wordMLPackage.getMainDocumentPart().getJaxbElement();

            org.docx4j.wml.Body body = doc.getBody();
            var contents = body.getContent();

            for (int index = 0; index < contents.size(); index++)
            {
                try
                {
                    P tempP = (P)contents.get(index);
                    if (tempP.toString().Contains(searchText))
                    {
                        p = tempP;
                        break;
                    }
                }
                catch
                {

                }
            }

            return p;
        }


        private static byte[] ReadImageFile(string urlImage)
        {
            File file = new File(urlImage);
            InputStream inputStream = null;
            byte[] bytes = null;

            try
            {
                inputStream = new java.io.FileInputStream(file);
                long fileLength = file.length();

                bytes = new byte[(int)fileLength];

                int offset = 0;
                int numRead = 0;

                while (offset < bytes.Length
                       && (numRead = inputStream.read(bytes, offset, bytes.Length - offset)) >= 0)
                {
                    offset += numRead;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                inputStream.close();
            }
            return bytes;
        }


        private static void CreateAnchorImageObject(WordprocessingMLPackage wordMLPackage, string urlImage, int xPos, int yPos, ref P p)
        {
            byte[] bytes = ReadImageFile(urlImage);

            String filenameHint = null;
            String altText = null;

            int id1 = 0;
            int id2 = 1;

            newImageToExistingPara(wordMLPackage, bytes, filenameHint, altText, id1, id2, ConvertInchesToEMU(xPos), ConvertInchesToEMU(yPos), ref p);
        }

        private static void CreateInlineImageObject(WordprocessingMLPackage wordMLPackage, string urlImage, ref P p)
        {
            byte[] bytes = ReadImageFile(urlImage);

            String filenameHint = null;
            String altText = null;

            int id1 = 0;
            int id2 = 1;

            newImageInlineToExistingPara(wordMLPackage, bytes, filenameHint, altText, id1, id2, ref p);
        }

        private static void newImageInlineToExistingPara(WordprocessingMLPackage wordMLPackage, byte[] bytes,
            String filenameHint, String altText, int id1, int id2,  ref P p)
        {
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
            Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2);

            org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();

            // Now add the inline in w:p/w:r/w:drawing
            org.docx4j.wml.Drawing drawing = Context.getWmlObjectFactory().createDrawing();
            org.docx4j.wml.R run = Context.getWmlObjectFactory().createR();

            p.getParagraphContent().add(run);

            run.getRunContent().add(drawing);
            drawing.getAnchorOrInline().add(inline);

        }


        private static void newImageToExistingPara(WordprocessingMLPackage wordMLPackage, byte[] bytes,
            String filenameHint, String altText, int id1, int id2, int xPos, int yPos, ref P p)
        {
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
            Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2);

            org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();

            string anchorXml = org.docx4j.XmlUtils.marshaltoString(inline, true, false, Context.jc, Namespaces.NS_WORD12, "anchor",
                    typeof(Inline));

            org.docx4j.dml.ObjectFactory dmlFactory = new org.docx4j.dml.ObjectFactory();
            org.docx4j.dml.wordprocessingDrawing.ObjectFactory wordDmlFactory = new org.docx4j.dml.wordprocessingDrawing.ObjectFactory();

            Anchor anchor = (Anchor)org.docx4j.XmlUtils.unmarshalString(anchorXml, Context.jc, typeof(Anchor));

            anchor.setBehindDoc(true);
            anchor.setSimplePos(dmlFactory.createCTPoint2D());

            anchor.setSimplePosAttr(new java.lang.Boolean(false));
            anchor.setPositionH(wordDmlFactory.createCTPosH());
            anchor.getPositionH().setPosOffset(new java.lang.Integer(xPos));
            anchor.getPositionH().setRelativeFrom(STRelFromH.MARGIN);
            anchor.setPositionV(wordDmlFactory.createCTPosV());
            anchor.getPositionV().setPosOffset(new java.lang.Integer(yPos));
            anchor.getPositionV().setRelativeFrom(STRelFromV.PAGE);
            anchor.setWrapNone(wordDmlFactory.createCTWrapNone());
            
            org.docx4j.wml.Drawing drawing = Context.getWmlObjectFactory().createDrawing();
            org.docx4j.wml.R run = Context.getWmlObjectFactory().createR();

            p.getParagraphContent().add(run);

            run.getRunContent().add(drawing);
            drawing.getAnchorOrInline().add(anchor);


        }

        private static int ConvertInchesToEMU(double inches)
        {
            double inchesToEMU = 914400;
            return Convert.ToInt32(inches * inchesToEMU);
        }

    }
}

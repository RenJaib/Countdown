using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Countdown
{
    public class WriteToWord
    {
        public static void CreateDocument(string docFilePath, string term, byte[] image)
        {
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application(); //requires to add Microsoft Office reference
            Microsoft.Office.Interop.Word._Document oDoc;
            oDoc = oWord.Documents.Add();
            
            //To see whats going on while populating the word document set Visible = true
            oWord.Visible = true;

            //Insert text
            Object oMissing = System.Reflection.Missing.Value;
            var oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = term;
            oPara1.Range.InsertParagraphAfter();

            //Here is the trick to insert a picture from a byte array into MS Word you need to 
            //convert the ByteArray into an Image and using the Clipboard paste it into the document
            Image sparePicture = ByteArrayToImage(image);
            Clipboard.SetDataObject(sparePicture); //requires to add reference system.windows.forms //copies the image similar to cntrl+c
            var oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara2.Range.Paste(); //paste the image
            oPara2.Range.InsertParagraphAfter();
            oDoc.SaveAs(docFilePath); //save to new document
            oDoc.Close();
            oWord.Quit();
        }

        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream ms = new MemoryStream(byteArrayIn))
            {
                Image returnImage = Image.FromStream(ms);
                return returnImage;
            }
        }

        //Method I use to test loading images from disk into byte[]'s and inserting them into word
        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            byte[] result = null;
            using (MemoryStream ms = new MemoryStream())
            {
                imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                result = ms.ToArray();
            }
            return result;
        }
    }
}

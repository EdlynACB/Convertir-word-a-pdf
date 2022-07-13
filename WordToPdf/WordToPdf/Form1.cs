using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\edlyn.castro\Desktop\Carta\Prueba2.docx";
            string pathpdf = @"C:\Users\edlyn.castro\Desktop\Carta\Prueba2.pdf";
            //string path = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).doc";
            //string pathpdf = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).pdf";
            try
            {
                System.IO.FileStream fs;
                fs = System.IO.File.Open(path, System.IO.FileMode.Open);
                var data = new byte[fs.Length];
                fs.Read(data, 0, Convert.ToInt32(fs.Length));
                fs.Close();


                FileStream fs2 = new FileStream(pathpdf, FileMode.OpenOrCreate);
                fs2.Write(data, 0, data.Length);
                fs2.Close();

                System.IO.File.WriteAllBytes(pathpdf, data);

                textBox1.Text = "Finalizo";
            }
            catch (Exception ex)
            {
                textBox1.Text = ex.Message;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).doc";
            string pathpdf = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).pdf";
            //string path = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).doc";
            //string pathpdf = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).pdf";
            try
            {
                //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

                //object oMissing = System.Reflection.Missing.Value;

                //word.Visible = false;
                //word.ScreenUpdating = false;

                //Object filename = (Object)path;

                //Document doc = word.Documents.Open(ref filename, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //doc.Activate();

                ////object outputFileName = wordFile.FullName.Replace(".doc", ".pdf");
                //object outputFileName = pathpdf;
                //object fileFormat = WdSaveFormat.wdFormatPDF;

                //// Save document into PDF Format
                //doc.SaveAs(ref outputFileName,
                //    ref fileFormat, ref oMissing, ref oMissing,
                //    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                //    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                //    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //// Close the Word document, but leave the Word application open.
                //// doc has to be cast to type _Document so that it will find the
                //// correct Close method.                
                //object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                //((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                //doc = null;

                //// word has to be cast to type _Application so that it will find
                //// the correct Quit method.
                //((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                //word = null;


        //        Dim word As Application = New Application()
        //Dim doc As Document = word.Documents.Open("c:\document.docx")
        //doc.Activate()
        //doc.SaveAs2("c:\document.pdf", WdSaveFormat.wdFormatPDF)
        //doc.Close()

                textBox1.Text = "Finalizo";
            }
            catch (Exception ex)
            {
                textBox1.Text = ex.Message;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Create a byte array that will eventually hold our final PDF
            string path = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).mht";
            string pathpdf = @"C:\Users\edlyn.castro\Desktop\Carta\Carta Consular (Titular).pdf";
            Byte[] bytes;
            try 
            {
                
                //Boilerplate iTextSharp setup here
                //Create a stream that we can write to, in this case a MemoryStream
                using (var ms = new MemoryStream())
                {

                    //Create an iTextSharp Document which is an abstraction of a PDF but **NOT** a PDF
                    using (var doc = new iTextSharp.text.Document())
                    {

                        //Create a writer that's bound to our PDF abstraction and our stream
                        using (var writer = PdfWriter.GetInstance(doc, ms))
                        {

                            writer.CloseStream = false;

                            //Open the document for writing
                            doc.Open();

                            //Our sample HTML and CSS
                            FileInfo fs = new FileInfo(path);
                            var example_html = fs.OpenText().ReadToEnd();
                            //var example_html = fs.Open(FileMode.Open, FileAccess.Read);

                            //string example_html = File.ReadAllText(path);                            
                            
                            example_html = @"<p>This <em>is </em><span class=""headline"" style=""text-decoration: underline;"">some</span> <strong>sample <em> text</em></strong><span style=""color: red;"">!!!</span></p>";
                            var example_css = @".headline{font-size:200%}";
                            

                            /**************************************************
                             * Example #1                                     *
                             *                                                *
                             * Use the built-in HTMLWorker to parse the HTML. *
                             * Only inline CSS is supported.                  *
                             * ************************************************/

                            //Create a new HTMLWorker bound to our document
                            using (var htmlWorker = new iTextSharp.text.html.simpleparser.HTMLWorker(doc))
                            {

                                //HTMLWorker doesn't read a string directly but instead needs a TextReader (which StringReader subclasses)
                                using (var sr = new StringReader(example_html))
                                {

                                    //Parse the HTML
                                    htmlWorker.Parse(sr);
                                }
                            }

                            /**************************************************
                             * Example #2                                     *
                             *                                                *
                             * Use the XMLWorker to parse the HTML.           *
                             * Only inline CSS and absolutely linked          *
                             * CSS is supported                               *
                             * ************************************************/

                            //XMLWorker also reads from a TextReader and not directly from a string
                            using (var srHtml = new StringReader(example_html))
                            {

                                //Parse the HTML
                                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, srHtml);

                            }

                            /**************************************************
                             * Example #3                                     *
                             *                                                *
                             * Use the XMLWorker to parse HTML and CSS        *
                             * ************************************************/

                            //In order to read CSS as a string we need to switch to a different constructor
                            //that takes Streams instead of TextReaders.
                            //Below we convert the strings into UTF8 byte array and wrap those in MemoryStreams
                            using (var msCss = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(example_css)))
                            {
                                using (var msHtml = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(example_html)))
                                {

                                    //Parse the HTML
                                    iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msHtml, msCss);
                                }
                            }


                            doc.Close();
                        }
                    }

                    //After all of the PDF "stuff" above is done and closed but **before** we
                    //close the MemoryStream, grab all of the active bytes from the stream
                    bytes = ms.ToArray();
                }

                //Now we just need to do something with those bytes.
                //Here I'm writing them to disk but if you were in ASP.Net you might Response.BinaryWrite() them.
                //You could also write the bytes to a database in a varbinary() column (but please don't) or you
                //could pass them to another function for further PDF processing.
                var testFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.pdf");
                System.IO.File.WriteAllBytes(testFile, bytes);

                //________________________________________

                //string strHtml = string.Empty;
                ////HTML File path -http://aspnettutorialonline.blogspot.com/
                //string htmlFileName = Server.MapPath("~") + "\\files\\" + "ConvertHTMLToPDF.htm";
                ////pdf file path. -http://aspnettutorialonline.blogspot.com/
                //string pdfFileName = Request.PhysicalApplicationPath + "\\files\\" + "ConvertHTMLToPDF.pdf";

                ////reading html code from html file
                //FileStream fsHTMLDocument = new FileStream(htmlFileName, FileMode.Open, FileAccess.Read);
                //StreamReader srHTMLDocument = new StreamReader(fsHTMLDocument);
                //strHtml = srHTMLDocument.ReadToEnd();
                //srHTMLDocument.Close();

                //strHtml = strHtml.Replace("\r\n", "");
                //strHtml = strHtml.Replace("\0", "");

                //CreatePDFFromHTMLFile(strHtml, pdfFileName);

               
                //Response.Write("pdf creation successfully with password -http://aspnettutorialonline.blogspot.com/");

            textBox1.Text = "Finalizo";
            }
            catch (Exception ex)
            {
                textBox1.Text = ex.Message;
            }
        }


        public void CreatePDFFromHTMLFile(string HtmlStream, string FileName)
        {
            try
            {
                //object TargetFile = FileName;
                //string ModifiedFileName = string.Empty;
                //string FinalFileName = string.Empty;

                ///* To add a Password to PDF -http://aspnettutorialonline.blogspot.com/ */
                //TestPDF.HtmlToPdfBuilder builder = new TestPDF.HtmlToPdfBuilder(iTextSharp.text.PageSize.A4);
                //TestPDF.HtmlPdfPage first = builder.AddPage();
                //first.AppendHtml(HtmlStream);
                //byte[] file = builder.RenderPdf();
                //File.WriteAllBytes(TargetFile.ToString(), file);

                //iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(TargetFile.ToString());
                //ModifiedFileName = TargetFile.ToString();
                //ModifiedFileName = ModifiedFileName.Insert(ModifiedFileName.Length - 4, "1");

                //string password = "password";
                //iTextSharp.text.pdf.PdfEncryptor.Encrypt(reader, new FileStream(ModifiedFileName, FileMode.Append), iTextSharp.text.pdf.PdfWriter.STRENGTH128BITS, password, "", iTextSharp.text.pdf.PdfWriter.AllowPrinting);
                ////http://aspnettutorialonline.blogspot.com/
                //reader.Close();
                //if (File.Exists(TargetFile.ToString()))
                //    File.Delete(TargetFile.ToString());
                //FinalFileName = ModifiedFileName.Remove(ModifiedFileName.Length - 5, 1);
                //File.Copy(ModifiedFileName, FinalFileName);
                //if (File.Exists(ModifiedFileName))
                //    File.Delete(ModifiedFileName);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}

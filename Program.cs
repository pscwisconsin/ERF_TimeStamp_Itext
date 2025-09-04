using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using System.Text;
using System.Net.Mail;
using System.Threading;


using Paragraph = iText.Layout.Element.Paragraph;
using Rectangle = iText.Kernel.Geom.Rectangle;
using System.Runtime.Remoting.Messaging;

using iText.Kernel.Pdf;
using iText.Layout.Properties;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;

namespace ERFTimestampiText
{
    internal class Program
    {
        

        private static string SMTP_HOST = "mail.testmail.com";
        private static string cnnPATH = ConfigurationManager.ConnectionStrings["cnnPath"].ConnectionString;
        private static string FilePath = ConfigurationManager.AppSettings["FilePath"];

        private static string LogFile = "";
        private static bool FirstWrite = true;

        private static string MYwriterFilePathName;
        private static string MYfileTablePathName;

        // =====================================================================================================================================
        // 
        // =====================================================================================================================================
        static void Main(string[] args)
        {


            if (FilePath == string.Empty)
            {
                FilePath = AppDomain.CurrentDomain.BaseDirectory;
            }


            LogFile = FilePath + "BatchLog.txt";

            WriteLog(DateTime.Now + " <= BEGIN TIMESTAMP PROCESS =>");

            string sqlStatement = "SELECT documentId, confidentialind, recvdt, docdesctxt, docsuubmituserId"
                         + ", file_stream.GetFileNamespacePath(1) AS filenamepath, b.name, b.stream_id"
                         + " FROM Pdocument a "
                         + " INNER JOIN pdocumentfiletable b ON a.streamid = b.streamid"
                         + " WHERE docfiletypecd = 'PDF' AND timestampind = 'N' AND timestamperrorind = 'N'"
                         + " ORDER BY documentid";
            DataTable dt = new DataTable();

            using (SqlConnection con = new SqlConnection(cnnPATH))
            {
                using (SqlCommand cmd = new SqlCommand(sqlStatement, con))
                {
                    cmd.CommandType = CommandType.Text;
                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(reader);
                }
            }
            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("there are documents...");

                foreach (DataRow row in dt.Rows)
                {
                    int doc_id = Convert.ToInt32(row["documentId"]);
                    string doc_recv_dt = Convert.ToString(row["recvdt"]);
                    string conf_ind = Convert.ToString(row["confidentialind"]);
                    //string file_name_path = Convert.ToString(row["file_name_path"]);
                    //string filetable_name = Convert.ToString(row["name"]);
                    string stream_id = Convert.ToString(row["streamid"]);
                    string doc_desc_txt = Convert.ToString(row["docdesctxt"]);
                    string doc_submit_logon_id = Convert.ToString(row["docsuubmituserId"]);

                    // original uploaded file (input file)
                    MYfileTablePathName = Convert.ToString(row["filenamepath"]);
                    // output file
                    MYwriterFilePathName = FilePath + DateTime.Now.ToString("yyyy-MM-dd hhmmss") + ".PDF";

                    WriteLog(DateTime.Now + " => START PROCESS => doc_id: " + doc_id);

                    // STEP 1 (will detected not pdf format and popup password protected)
                    if (!checkUploadedFile())
                    {
                        Console.WriteLine("NOT FOUND PDF file or file size is Zero, do someting");
                        string message = "<p><b>UPLOADED DOCUMENT: PDF NO CONTENT</b></p>"
                                        + "<br /><b>doc ID: </b>" + doc_id + "</p>"
                                        + "<br /><b>Description: </b>" + doc_desc_txt + "</p>"
                                        + "<br /><b>Submitted by: </b>" + doc_submit_logon_id + "</p>"
                                        ;
                        EmailRM(message);

                        UpdateTimestampIssue(doc_id, "File size 0");

                        continue;
                    }

                    // STEP 2 (will detected not pdf format and popup password protected)
                    if (!isPDF())
                    {
                        Console.WriteLine("NOT PDF file, do someting");
                        string message = "<p><b>UPLOADED DOCUMENT: CORRUPTED FILE</b></p>"
                                        + "<br /><b>doc ID: </b>" + doc_id + "</p>"
                                        + "<br /><b>Description: </b>" + doc_desc_txt + "</p>"
                                        + "<br /><b>Submitted by: </b>" + doc_submit_logon_id + "</p>"
                                        ;
                        EmailRM(message);

                        UpdateTimestampIssue(doc_id, "Corrupted File or Popup Window");

                        continue;
                    }

                    // STEP 2.2 (will detected if file cannot open)
                    // 02/13/2023 Gaysorn - per story#7116 
                    if (isCorruptedPDF())
                    {
                        Console.WriteLine("CANNOT OPEN PDF file, do someting");
                        string message = "<p><b>UPLOADED DOCUMENT: CORRUPTED FILE CANNOT OPEN</b></p>"
                                        + "<br /><b>doc ID: </b>" + doc_id + "</p>"
                                        + "<br /><b>Description: </b>" + doc_desc_txt + "</p>"
                                        + "<br /><b>Submitted by: </b>" + doc_submit_logon_id + "</p>"
                                        ;
                        //EmailRM(message);

                        UpdateTimestampIssue(doc_id, "Corrupted File");

                        continue;

                    }


                    // STEP 3 (without popup password protected)
                    if (isPassword())
                    {
                        File.Delete(MYwriterFilePathName);
                        Console.WriteLine("Password Protected, do something");
                        string message = "<p><b>UPLOADED DOCUMENT: PASSWORD PROTECTED</b></p>"
                                        + "<p><b>doc ID: </b>" + doc_id + "</p>"
                                        + "<p><b>Description: </b>" + doc_desc_txt + "</p>"
                                        + "<p><b>Submitted by: </b>" + doc_submit_logon_id + "</p>"
                                        ;
                        EmailRM(message);

                        UpdateTimestampIssue(doc_id, "Password Protected");

                        continue;
                    }

                    // STEP 3
                    if (!StampDocument(confidentialind, streamid, documentId, recvdt)){

                        Console.WriteLine("cannot stampDocument, do something");
                        continue;
                    }

                    WriteLog("Finished Stamped at: " + DateTime.Now);
                }
            }
            else
            {
                WriteLog(DateTime.Now + " No Uploaded Document to stamp");
            }
        }

        // ===============================================================
        //
        // ===============================================================
        private static bool checkUploadedFile()
        {
            WriteLog("Check: checkUploadedFile");
            FileInfo fInfo = new FileInfo(MYfileTablePathName);
            if (fInfo.Exists)
            {
                if (fInfo.Length == 0)
                {
                    // IT should get notify
                    WriteLog("PDF file has size of 0.");
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                // IT should get notify
                WriteLog("PDF not found in FileTable.");
                return false;
            }
        }

        // ===============================================================
        // check if pdf file
        // ===============================================================
        private static bool isPDF()
        {
            WriteLog("Check: isPDF");
            try
            {
                PdfDocument pdf = new PdfDocument(new PdfReader(MYfileTablePathName));
                PdfAConformanceLevel level = pdf.GetReader().GetPdfAConformanceLevel();
                pdf.Close();
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("ERROR isPDF:" + ex.Message);
                return false;
            }

        }

        // ===============================================================
        // 02/13/2023  - check if pdf can open (corrupted)
        // ===============================================================
        private static bool isCorruptedPDF()
        {
            WriteLog("Check: CorruptedPDF");
            try
            {
                PdfDocument pdf = new PdfDocument(new PdfReader(MYfileTablePathName));
                pdf.GetFirstPage();
                pdf.Close();
                return false;
            }
            catch (Exception ex)
            {
                WriteLog("ERROR corruptedPDF:" + ex.Message);
                return true;
            }
        }

        // ===============================================================
        //
        // ===============================================================
        private static bool isPassword()
        {
            WriteLog("Check: isPassword");
            using (PdfReader myReader = new PdfReader(MYfileTablePathName))
            using (PdfWriter myWriter = new PdfWriter(MYwriterFilePathName))
            {
                try
                {
                    PdfDocument myPDF = new PdfDocument(myReader, myWriter);
                    return false;
                }
                catch (Exception ex)
                {
                    WriteLog("ERROR isPassword: " + ex.Message);
                    return true;
                }
            }
        }

        private static bool StampDocument(string confidentialFile, string streamId, int documentId, string receivedDateTime)
        {
            WriteLog("Stamp Document");
            // use high-level object (paragraph) to add stamp
            //(1) PSC REF#: on top of first page
            Paragraph stampREF = new Paragraph("REF#:" + documentId)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetFont(PdfFontFactory.CreateFont(StandardFonts.COURIER))
            .SetFontColor(ColorConstants.BLUE)
            .SetFontSize(16);
            
            //(2) 
            Paragraph stampPSC = new Paragraph("Organization title..Insert Here")
            .SetTextAlignment(TextAlignment.CENTER)
            .SetFont(PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD))
            .SetFontColor(ColorConstants.BLUE)
            .SetFontSize(9)
            .SetRotationAngle((Math.PI / 180) * -90);
            stampPSC.Add("\n");
            stampPSC.Add("RECEIVED: " + receivedDateTime);
            // -----------------------------------------------------------
            // show border (debug) if enable set rectangle2(x, y, 27, 50)
            // -----------------------------------------------------------
            //stampPSC.SetBorder(new SolidBorder(1));

            // (1)original uploaded file (input class)
            // (SetUnethicalReading(true) to inore the Owner password (ex changing document not allowed)
            PdfReader myReader = new PdfReader(MYfileTablePathName);
            //myReader.SetUnethicalReading(true);

            // (2)create an instance of PdfWriter (output class) to specify the target file name & path
            PdfWriter myWriter = new PdfWriter(MYwriterFilePathName);

            // (3)initialize PDF document (manipulate the original file)
            //create a PdfDocument object using the reader and the writer object as parameters
            PdfDocument myPDF = new PdfDocument(myReader, myWriter);

            // get the first page with the page size
            PdfPage page = myPDF.GetFirstPage();

            var pageSize = page.GetPageSizeWithRotation();
            Rectangle cropBox = page.GetCropBox();
            Rectangle mediaBox = page.GetMediaBox();

            // must use PdfCanvas to modify 
            PdfCanvas pdfCanvas = new PdfCanvas(page);

            try
            {
                // ******************************************************************************************************
                // PORTRAIT PDF size: 612,792
                // LANDSCAPE PDF size: 792,612
                // position text rectangle(x-from left to right, y-from bottom to top, xx-how width, yy-how height)
                // ******************************************************************************************************
                // --------------------------------------------------------------------
                // PSC REF#:, set space 182 from top page, space 102 from right page
                // --------------------------------------------------------------------
                //float x = pageSize.GetRight() - 182;
                //float y = pageSize.GetTop() - 102;
                switch (page.GetRotation())
                {
                    case 0:
                        float x = mediaBox.GetRight() - 182;
                        float y = mediaBox.GetTop() - 102;
                        Rectangle rectangleREF = new Rectangle(x, y, 150, 100);
                        Canvas canvas = new Canvas(pdfCanvas, rectangleREF);
                        canvas.Add(stampREF);
                        canvas.Close();

                        float xx = mediaBox.GetRight() - 42;
                        float yy = mediaBox.GetTop() - 92;
                        Rectangle rectanglePSC = new Rectangle(xx, yy, 25, 50);
                        Canvas canvas2 = new Canvas(pdfCanvas, rectanglePSC);
                        canvas2.Add(stampPSC);
                        canvas2.Close();
                        break;
                    case 90:

                        x = 0;
                        y = mediaBox.GetTop() - 102;
                        rectangleREF = new Rectangle(x, y, 25, 25);
                        canvas = new Canvas(pdfCanvas, rectangleREF);
                        canvas.Add(stampREF.SetRotationAngle((Math.PI / 180) * 90));
                        canvas.Close();

                        xx = 92;
                        yy = mediaBox.GetTop() - 250;

                        rectanglePSC = new Rectangle(xx, yy,250,250);
                        canvas2 = new Canvas(pdfCanvas, rectanglePSC);
                        canvas2.Add(stampPSC.SetRotationAngle((Math.PI / 180) * 0));
                        canvas2.Close();
                        break;
                    case 180:
                        x = 42;
                        y = -66;
                        rectangleREF = new Rectangle(x, y, 150, 100);
                        canvas = new Canvas(pdfCanvas, rectangleREF);
                        canvas.Add(stampREF.SetRotationAngle((Math.PI / 180) * 180));
                        canvas.Close();

                        xx = 42;
                        yy = 260;

                        rectanglePSC = new Rectangle(xx, yy, 25, 50);
                        canvas2 = new Canvas(pdfCanvas, rectanglePSC);
                        canvas2.Add(stampPSC.SetRotationAngle((Math.PI / 180) * 90));
                        canvas2.Close();
                        break;
                    case 270:
                        x = mediaBox.GetRight() -45;
                        y = 200;
                        rectangleREF = new Rectangle(x, y, 25, 25);
                        canvas = new Canvas(pdfCanvas, rectangleREF);
                        canvas.Add(stampREF.SetRotationAngle((Math.PI / 180) * -90));
                        canvas.Close();

                        xx = mediaBox.GetRight() - 300;
                        yy = -180;

                        rectanglePSC = new Rectangle(xx, yy, 250, 250);
                        canvas2 = new Canvas(pdfCanvas, rectanglePSC);
                        canvas2.Add(stampPSC.SetRotationAngle((Math.PI / 180) * 180));
                        canvas2.Close();
                        break;
                    default:
                        break;
                }





                // ****************************************************************************************
                // HANDLE CONFIDENTIAL STAMP EVERY PAGES
                // ****************************************************************************************
                try
                {
                    //confidentialFile = "Y";
                    if (confidentialFile == "Y")
                    {
                        WriteLog("Stamp Confidential");
                        Paragraph stampConfidential = new Paragraph("CONFIDENTIAL")
                            .SetTextAlignment(TextAlignment.CENTER)
                            .SetFont(PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD))
                            .SetFontColor(ColorConstants.RED)
                            .SetFontSize(22)
                            .SetVerticalAlignment(VerticalAlignment.TOP);

                        float pageX = (pageSize.GetLeft() + pageSize.GetRight()) / 2;
                        float pageY = pageSize.GetTop() - 12;

                        // how many page??
                        int numberOfPage = myPDF.GetNumberOfPages();
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            Rectangle mediaBoxConf = myPDF.GetPage(i).GetMediaBox();
                            PdfCanvas myCanvas = new PdfCanvas(myPDF.GetPage(i));
                            
                            switch (myPDF.GetPage(i).GetRotation())
                            {
                                case 0:

                                    float x = mediaBoxConf.GetRight() / 2 - 66;
                                    float y = mediaBoxConf.GetTop() - 112;
                                    Rectangle rectangleREF = new Rectangle(x, y, 150, 100);
                                    Canvas canvasConfidentail = new Canvas(myCanvas, rectangleREF);
                                    canvasConfidentail.Add(stampConfidential.SetRotationAngle((Math.PI / 180) * 0));
                                    canvasConfidentail.Close();
                                    break;
                                case 90:
                                    x = 0;
                                    y = mediaBoxConf.GetTop() / 2;
                                    rectangleREF = new Rectangle(x, y, 150, 100);
                                    canvasConfidentail = new Canvas(myCanvas, rectangleREF);
                                    canvasConfidentail.Add(stampConfidential.SetRotationAngle((Math.PI / 180) * 90));
                                    canvasConfidentail.Close();

                                    break;
                                case 180:
                                    x = mediaBoxConf.GetRight() / 2 - 66;
                                    y = -66;
                                    rectangleREF = new Rectangle(x, y, 150, 100);
                                    canvasConfidentail = new Canvas(myCanvas, rectangleREF);
                                    canvasConfidentail.Add(stampConfidential.SetRotationAngle((Math.PI / 180) * 180));
                                    canvasConfidentail.Close();

                                    break;
                                case 270:

                                    x = mediaBoxConf.GetRight() - 50;
                                    y = mediaBoxConf.GetTop() / 2;
                                    rectangleREF = new Rectangle(x, y, 150, 100);
                                    canvasConfidentail = new Canvas(myCanvas, rectangleREF);
                                    canvasConfidentail.Add(stampConfidential.SetRotationAngle((Math.PI / 180) * -90));
                                    canvasConfidentail.Close();

                                    break;
                                default:
                                    break;
                            }

                        }
                    }
                }
                catch (Exception exConfidential)
                {
                    EmailIT("ERROR Exception StampDocument(Stamp Confidential):" + exConfidential.Message);
                    WriteLog("ERROR StampDocument-Confidential: " + exConfidential.Message);
                    return false;
                }
            }
            catch (Exception ex)
            {
                EmailIT("ERROR Exception StampDocument:" + ex.Message);
                WriteLog("ERROR StampDocument: " + ex.Message);
                return false;
            }
            finally
            {
                // must close reader before can save it
                myPDF.Close();
                myReader.Close();
            }

            if (SaveStampedDocument(streamId, MYwriterFilePathName))
            {
                // clean the temp file
                File.Delete(MYwriterFilePathName);
                return true;
            }
            else
            {
                // clean the temp file
                File.Delete(MYwriterFilePathName);
                return false;
            }
            
        }

        // =====================================================================================================================================
        // update FileTable with stamped document ane update the time_stamp_ind
        // =====================================================================================================================================
        private static bool SaveStampedDocument(string stream_id, string fileNamePath)
        {
            WriteLog("Update Database: Stamped Document ");

            byte[] fileContent;
            using (var fs = new FileStream(fileNamePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = new BinaryReader(fs))
                {
                    fileContent = reader.ReadBytes((int)fs.Length);
                }
            }

            int returnCode;
            string returnMessage;

            using (SqlConnection con = new SqlConnection(cnnPATH))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "batchtimestampUpdatefiletable";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Connection = con;
                    cmd.Parameters.Add("@stream_id", SqlDbType.NVarChar, 255).Value = streamid;
                    cmd.Parameters.AddWithValue("@DataContent", fileContent);
                    cmd.Parameters.Add("@ReturnMessage", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                    cmd.Parameters.Add("@ReturnCode", SqlDbType.Int).Direction = ParameterDirection.Output;

                    con.Open();
                    cmd.ExecuteNonQuery();
                    con.Close();

                    returnMessage = Convert.ToString(cmd.Parameters["@ReturnMessage"].Value);
                    returnCode = Convert.ToInt16(cmd.Parameters["@ReturnCode"].Value);
                    if (returnCode == 1)
                    {
                        return true;
                    }
                    else
                    {
                        EmailIT("Saved Failed filetable issues with message: " + returnMessage);  
                        return false;
                    }
                }
            }
        }

        // ==========================================================
        // email IT if having issues
        // ==========================================================
        private static void EmailIT(string ErrorMessage)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.Priority = MailPriority.Normal;
            mailMessage.IsBodyHtml = true;
            mailMessage.Subject = " TimeStamp Uploaded Document Issues (iText)";
            mailMessage.From = new MailAddress("noreply@testmail.com");
            mailMessage.To.Add(new MailAddress("errors@testmail.com"));
            mailMessage.Body = "Timestamp error, please check the following message:<p>" + ErrorMessage + "</p>";
            try
            {
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = SMTP_HOST;
                smtpClient.Send(mailMessage);

            }
            catch (Exception)
            { }

        }

        // ==========================================================
        // email record if having issues
      
        // ==========================================================
        private static void EmailRM(string ErrorMessage)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.Priority = MailPriority.Normal;
            mailMessage.IsBodyHtml = true;
            mailMessage.Subject = "Timestamp Uploaded Document Issues";
            mailMessage.From = new MailAddress("noreply@testmail.com");
            mailMessage.To.Add(new MailAddress("record@testmail.com"));
            mailMessage.Bcc.Add(new MailAddress("errors@testmail.com"));
            mailMessage.Body = "The following PDF file is either corrupted or password protected, please <b>REJECT</b> and ask for new PDF file.<p>" + ErrorMessage + "</p>";

            try
            {
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = SMTP_HOST;
                smtpClient.Send(mailMessage);

            }
            catch (Exception)
            { }

        }

        // ==========================================================
        //
        // ==========================================================
        private static void UpdateTimestampIssue(int doc_id, string issueMessage)
        {
            string sqlStatement = @"UPDATE TimeStampDatabase..Pdocument"
                                + " SET  timestamperrorind = 'Y', timestampissuedesc = @issueMessage"
                                + " WHERE documentid = @doc_id";

            using (SqlConnection con = new SqlConnection(cnnPATH))
            {
                using (SqlCommand cmd = new SqlCommand(sqlStatement, con))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@doc_id", documentid);
                    cmd.Parameters.AddWithValue("@issueMessage", issueMessage);
                    //con.Open();
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                }
            }

        }

        // ==========================================================
        // write log file
        // ==========================================================
        private static void WriteLog(string logMessage)
        {
            StreamWriter sw = null;
            if (FirstWrite)
            {
                FirstWrite = false;
                sw = File.CreateText(LogFile);
            }
            else
            {
                sw = new StreamWriter(LogFile, true);
            }
            sw.WriteLine(logMessage + "\r\n");
            sw.Close();
        }


        // ===============================================================
        // sample (NOT USE)
        // ===============================================================
        private static bool isPasswordProtected()
        {
            string tempFileNamePath = FilePath + DateTime.Now.ToString("yyyy-MM-dd hhmmss") + ".PDF";
            //PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(filePath));

            PdfReader myReader = new PdfReader(MYfileTablePathName);
            PdfDocument myPDF;

            try
            {
                //PdfReader myReader = new PdfReader(fileNamePath);

                // DO NOT STAMP WITH OWNER PASSWORD (if want, must have SetUnethicalReading(true)
                //myReader.SetUnethicalReading(true);               

                PdfWriter myWriter = new PdfWriter(tempFileNamePath);
                myPDF = new PdfDocument(myReader, myWriter);

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return true;
            }
            finally
            {
                myReader = null;
                myPDF = null;
                File.Delete(tempFileNamePath);

            }

            // this exception means the PDF cannot be opened at all
            //catch (iText.Kernel.Exceptions.BadPasswordException)
            //{
            //    return false;
            //}
        }

    }
}

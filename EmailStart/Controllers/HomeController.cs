using EmailStart.Models;
using OpenPop.Mime;
using OpenPop.Pop3;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace EmailStart.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHeadersRepository _headersRepository;

        public HomeController(IHeadersRepository headersRepository)
        {
            _headersRepository = headersRepository;
        }

        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            return View();
        }
        /*
         * 
         * Server name: outlook.office365.com
            Port: 995
            Encryption method: TLS
         * 
         * 
         * */

        [HttpPost]
        public List<Message> FetchAllMessages()
        {
            var hostname = "outlook.office365.com";
            var port = 995;
            var useSsl = true;
            var username = "xxxx";
            var password = "xxxx";
            {

                using (Pop3Client client = new Pop3Client())
                {
                    // Connect to the server
                    client.Connect(hostname, port, useSsl);

                    // Authenticate ourselves towards the server/
                    client.Authenticate(username, password);

                    // Get the number of messages in the inbox
                    int messageCount = client.GetMessageCount();

                    // We want to download all messages
                    List<Message> allMessages = new List<Message>(messageCount);

                    // Messages are numbered in the interval: [1, messageCount]
                    // Ergo: message numbers are 1-based.
                    // Most servers give the latest message the highest number
                    for (int i = messageCount; i > 0; i--)
                    {
                        var messageMain = client.GetMessage(i);
                        var messageHeaders = client.GetMessageHeaders(i);
                        var messagePart = client.GetMessage(i).MessagePart;
                        var attachments = client.GetMessage(i).FindAllAttachments();
                        var from = messageHeaders.From;
                        var subject = messageHeaders.Subject;
                        var date = messageHeaders.DateSent;
                        var messageId = messageHeaders.MessageId;

                        checkMessages(messageId);
                        if (ischecked == 0)
                        {

                            if (attachments != null || attachments.Count > 0)
                            {
                                //int asd = attachments.Count();
                                //for (i = 0; i == asd; i++)
                                //{
                                foreach (var item in attachments)
                                {
                                    string fileName = item.FileName;

                                    if (fileName == "(no name)" || !fileName.Contains("."))
                                    {
                                        continue;
                                    }

                                    //string extension = item.FileName.Split('.')[1];
                                    string extension = fileName.Substring(fileName.LastIndexOf(".") + 1);
                                    Extension = extension;

                                    if (extension == "png" || extension == "jpeg" || extension == "jpg" || extension == "gif" || extension == "html" || extension == "pdf")
                                    {
                                        continue;
                                    }

                                    System.IO.File.WriteAllBytes(@"C:\Users\ozgen.akincilar\Desktop\MailOutput\Dummy." + extension, item.Body);


                                    checkMessages(messageId);
                                    if (ischecked == 0)
                                    {
                                        Subject = subject;
                                        Mahmut(messageId);
                                    }



                                    if (extension == "txt")
                                    {
                                        string[] cc = System.IO.File.ReadAllLines(@"C:\Users\ozgen.akincilar\Desktop\MailOutput\Dummy." + extension, Encoding.Default);
                                        int lenghtTxt = cc.GetLength(0);
                                        for (int ii = 0; ii < lenghtTxt; ii++)
                                        {
                                            string aa = cc[ii];
                                            _headersRepository.Mt940txtKaydet(messageId, aa);
                                        }
                                    }
                                    if (extension == "xlsx" || extension=="xls")
                                    {
                                        HeadersModel hm = new HeadersModel();
                                        hm.From = from.ToString();
                                        hm.Subject = subject;
                                        hm.FileName = fileName;
                                        hm.SentDate = date;
                                        hm.MessageId = messageId;
                                        _headersRepository.UstBilgiKaydet(hm);

                                    }
                                    else break;
                                    //if (statusChecked == 0)
                                    //{
                                    //    Subject = subject;
                                    //}
                                    //else break;

                                }
                                //}

                            }
                        }
                        else break;
                        allMessages.Add(client.GetMessage(i));
                    }

                    // Now return the fetched messages
                    return allMessages;
                }
            }
        }


        [HttpPost]
        public static Message SaveAndLoadFullMessage(Message message)
        {
            // FileInfo about the location to save/load message
            FileInfo file = new FileInfo(@"C:\Users\ozgen.akincilar\Desktop\MailOutput\someFile.eml");

            // Save the full message to some file
            message.Save(file);

            // Now load the message again. This could be done at a later point
            Message loadedMessage = Message.Load(file);

            // use the message again
            return loadedMessage;
        }

        string Subject = "";
        string Extension = "";
        [HttpPost]
        public void Mahmut(string MessageId)//
        {
            string filePath = "";
            if (Extension == "xls")
            {
                 filePath = @"C:\Users\ozgen.akincilar\Desktop\MailOutput\Dummy.xls";
            }
            else
                 filePath = @"C:\Users\ozgen.akincilar\Desktop\MailOutput\Dummy.xlsx";

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            var missing = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //Here we say which tabs going to insert into database
            
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange; //Get only worked range in excel sheet
            Array myValues = (Array)xlRange.Cells.Value2;
            List<BelgeModel> zartingen = new List<BelgeModel>();
          
            for (int i = 1; i <= myValues.GetLength(0); i++)
            {
                BelgeModel cart = new BelgeModel();
                List<string> kurt = new List<string>();
                int asd = i;

                for (int j = 1; j <= myValues.GetLength(1); j++)
                {
                    string zart = "";

                    if (myValues.GetValue(i, j) == null)
                    {
                        zart = "";
                    }
                    else
                        zart = myValues.GetValue(i, j).ToString();
                    kurt.Add(zart);
                    #region zurt
                    //if (myValues.GetValue(i, j + 1) == null)
                    //{
                    //    zart2 = "Null";
                    //}
                    //else
                    //    zart2 = myValues.GetValue(i, j + 1).ToString();

                    //if (myValues.GetValue(i, j + 2) == null)
                    //{
                    //    zart3 = "Null";
                    //}
                    //else
                    //    zart3 = myValues.GetValue(i, j + 2).ToString();

                    //if (myValues.GetValue(i, j + 3) == null)
                    //{
                    //    zart4 = "Null";
                    //}
                    //else
                    //    zart4 = myValues.GetValue(i, j + 3).ToString();

                    //if (myValues.GetValue(i, j + 4) == null)
                    //{
                    //    zart5 = "Null";
                    //}
                    //else
                    //    zart5 = myValues.GetValue(i, j + 4).ToString();

                    //if (myValues.GetValue(i, j + 5) == null)
                    //{
                    //    zart6 = "Null";
                    //}
                    //else
                    //    zart6 = myValues.GetValue(i, j + 5).ToString();

                    //if (myValues.GetValue(i, j + 6) == null)
                    //{
                    //    zart7 = "Null";
                    //}
                    //else
                    //    zart7 = myValues.GetValue(i, j + 6).ToString();

                    //if (myValues.GetValue(i, j + 7) == null)
                    //{
                    //    zart8 = "Null";
                    //}
                    //else
                    //    zart8 = myValues.GetValue(i, j + 7).ToString();

                    //if (myValues.GetValue(i, j + 8) == null)
                    //{
                    //    zart9 = "Null";
                    //}
                    //else
                    //    zart9 = myValues.GetValue(i, j + 8).ToString();
                    //if (myValues.GetValue(i, j + 9) == null)
                    //{
                    //    zart10 = "Null";
                    //}
                    //else
                    //    zart10 = myValues.GetValue(i, j + 9).ToString();


                    //string zart2 = myValues.GetValue(i, j + 1).ToString();
                    //string zart3 = myValues.GetValue(i, j + 2).ToString();
                    //string zart4 = myValues.GetValue(i, j + 3).ToString();
                    //string zart5 = myValues.GetValue(i, j + 4).ToString();
                    //string zart6 = myValues.GetValue(i, j + 5).ToString();
                    //string zart7 = myValues.GetValue(i, j + 6).ToString();
                    //string zart8 = myValues.GetValue(i, j + 7).ToString();
                    //string zart9 = myValues.GetValue(i, j + 8).ToString();
                    //string zart10 = myValues.GetValue(i, j + 9).ToString();
                    //string MessageId = "asdasdasd"; 
                    #endregion
                }
                cart.Kolon = kurt;
                zartingen.Add(cart);
                var asqd = i;
            }
            _headersRepository.ZikKaydet(zartingen, MessageId);
            if (Subject == "Fark Raporu")
            {
                _headersRepository.FarkRaporuKaydet(zartingen, MessageId);
            }
            else if (Subject == "Stok")
            {
                _headersRepository.StokKaydet(zartingen, MessageId);
            }
            else if (Subject == "iade" || Subject == "İade" || Subject == "Iade")
            {
                _headersRepository.IadeKaydet(zartingen, MessageId);
            }
            else if(Subject== "RE: kpı")
            {


                _headersRepository.KpiKaydet(zartingen, MessageId);

            }

        }

        int ischecked;
        public int checkMessages(string MessageId)
        {
            ischecked = _headersRepository.checkMessage(MessageId);
            return ischecked;
        }

        int statusChecked;
        public int checkStatus(string MessageId)
        {
            statusChecked = _headersRepository.checkStatu(MessageId);
            return statusChecked;
        }

    }



    //class Message { }


}

using Spire.Barcode;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GenerateQRCode
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        public byte[] createQrCode(string s_fileNAme_Only = "")
        {
            BarcodeSettings settings = new BarcodeSettings();
            settings.Type = BarCodeType.QRCode;

            string[] sUrl = { };

            s_fileNAme_Only = s_fileNAme_Only.Replace("_", "&");
            sUrl = s_fileNAme_Only.Split('&');

            for (int i = 0; i < sUrl.Length; i++)  // 배열은 0 부터 저장되며, 배열의 길이만큼 순환
            {                
                if (i == sUrl.Length-1) 
                {
                    // 마지막 빈칸이 발견되면 빈칸은 전까지 값 가져옴.
                    sUrl[i] = sUrl[i].Substring(0, sUrl[i].LastIndexOf(' '));
                }
                //Console.WriteLine(i + "번째 배열 ==> " + sUrl[i]);
            }

            string value = $@"https://e-checking.kr.sgs.com/construction/default.aspx?ReportNo=" + sUrl[0] + "&IssueDate=" + sUrl[1] + "&RandomNo=" + sUrl[2];            

            settings.Data = value;
            settings.Data2D = value;

            settings.QRCodeDataMode = QRCodeDataMode.AlphaNumber;
            settings.QRCodeECL = QRCodeECL.H;
            
            settings.ShowText = false;
            settings.AutoResize = false;
            settings.ImageWidth = 166.6f;
            settings.ImageHeight = 130.5f;
            
            settings.X = 2.4f;
            settings.Y = 2.4f;
            settings.DpiX = 96;
            settings.DpiY = 96;

            BarCodeGenerator generator = new BarCodeGenerator(settings);
            Image image = generator.GenerateImage();

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                return ms.ToArray();
            }
        }

        private void replaceQr(Document word, byte[] samplePicture, string s_filepath = "")
        {
            string rpt = s_filepath;
            string s_Paht = Path.GetDirectoryName(rpt);
            string s_NAme = Path.GetFileNameWithoutExtension(rpt) + ".docx";
            string s_NAme_pdf = Path.GetFileNameWithoutExtension(rpt) + ".pdf";
            string s_full = s_Paht + "\\" + s_NAme;
            string s_full_pdf = s_Paht + "\\" + s_NAme_pdf;

            Section section = word.Sections[0];
            HeaderFooter header = section.HeadersFooters.FirstPageFooter;

            //Get the first table
            Table table = header.Tables[1] as Table;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    //Foreach paragraph in the cell
                    foreach (Paragraph par in table.Rows[i].Cells[j].Paragraphs)
                    {
                        //Get each document object of paragraph items
                        foreach (DocumentObject docObject in par.ChildObjects)
                        {
                            //If Type of Document Object is Picture, add it into image list
                            if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                DocPicture picture = docObject as DocPicture;

                                if (picture.AlternativeText == "IS_QRCode")
                                {
                                    float w = picture.Width;
                                    float h = picture.Height;
                                    picture.LoadImage(samplePicture);

                                    picture.Width = w;
                                    picture.Height = h;
                                }
                            }
                        }
                    }
                }
            }

            word.SaveToFile(s_full_pdf); // pdf 로 다시 저장
            
            //doc 파일 삭제
            if (File.Exists(s_full))
            {
                try
                {
                    File.Delete(s_full);
                }
                catch (Exception e)
                {
                    Console.WriteLine("The deletion failed: {0}", e.Message);
                }
            }

            //word.SaveToFile(@"C:\Users\hiel_kim\OneDrive - SGS\Documents\All\C#\KR\Dongtan\E&E\test1.docx");    // word 새롭게 저장
            word.Close();

            //TextSelection[] selections = word.FindAllString("IS_QRCode", true, true);
            //int index = 0;
            //TextRange range = null;

            //foreach (TextSelection selection in selections)
            //{
            //    DocPicture pic = new DocPicture(word);

            //    //float w = pic.Width;
            //    //float h = pic.Height;

            //    pic.LoadImage(samplePicture);

            //    pic.Width = 60;
            //    pic.Height = 80;

            //    range = selection.GetAsOneRange();
            //    index = range.OwnerParagraph.ChildObjects.IndexOf(range);
            //    range.OwnerParagraph.ChildObjects.Insert(index, pic);
            //    range.OwnerParagraph.ChildObjects.Remove(range);
            //}

            //Get Each Section of Document
            //foreach (Section section in word.Sections)
            //{
            //    //Loop through the paragraphs of the section
            //    foreach (DocumentObject docObj in section.HeadersFooters)
            //    {
            //        //Loop through the child elements of paragraph
            //        foreach (DocumentObject docObjChild in docObj.ChildObjects)
            //        {
            //            //If Type of Document Object is Picture, Extract.
            //            if (docObjChild.DocumentObjectType == DocumentObjectType.Paragraph)
            //            {
            //                //Get Each Paragraph of Section
            //                foreach (Paragraph paragraph in docObjChild.ChildObjects)
            //                {
            //                    //Get Each Document Object of Paragraph Items
            //                    foreach (DocumentObject docObject4 in paragraph.ChildObjects)
            //                    {
            //                        //If Type of Document Object is Picture, Extract.
            //                        if (docObject4.DocumentObjectType == DocumentObjectType.Picture)
            //                        {
            //                            DocPicture picture = docObject4 as DocPicture;
            //                            if (picture.Title == "IS_QRCode")
            //                            {
            //                                //Replace the image
            //                                picture.LoadImage(samplePicture);
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            //Loop through the paragraphs of the section
            //foreach (Paragraph foot in word.Sections[0].Paragraphs)
            //{
            //    //Loop through the child elements of paragraph
            //    foreach (DocumentObject docObj in foot.ChildObjects)
            //    {
            //        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            //        {
            //            DocPicture picture = docObj as DocPicture;
            //            if (picture.AlternativeText == "IS_QRCode")
            //            {
            //                //Replace the image
            //                picture.LoadImage(samplePicture);
            //            }
            //        }
            //    }
            //}

            //Loop through the paragraphs of the section
            //foreach (Paragraph foot in header.Paragraphs)
            //{
            //    //Loop through the child elements of paragraph
            //    foreach (DocumentObject docObj in foot.ChildObjects)
            //    {
            //        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            //        {
            //            DocPicture picture = docObj as DocPicture;
            //            if (picture.AlternativeText == "IS_QRCode")
            //            {
            //                //Replace the image
            //                picture.LoadImage(samplePicture);
            //            }
            //        }
            //    }
            //}

            //foreach (Table tb in header.Tables) { }

            //Get the first section
            //Section sec = word.Sections[0];

            //Create image list to collect extracted image
            //List<Image> images = new List<Image>();            

            //Document doc = new Document();
            //doc.LoadFromFile(@"Template.docx");

            //Image newImage = Image.FromFile(@"E-iceblue.png");
            //DocPicture newPic = new DocPicture(word);
            //newPic.LoadImage(samplePicture);

            //Section section = word.Sections[0];
            //HeaderFooter header = section.HeadersFooters.FirstPageFooter;

            //foreach (Table tb in header.Tables)
            //{
            //    List<DocumentObject> pictures = new List<DocumentObject>();

            //    foreach (DocumentObject docObj in tb.para)
            //    { 

            //    }

            //    //Get all pictures in the header
            //    foreach (DocumentObject docObj in tb.ChildObjects)
            //    {
            //        if (docObj.DocumentObjectType == DocumentObjectType.TableRow)
            //        {
            //            foreach (DocumentObject tbr in docObj.ChildObjects)
            //            {
            //                MessageBox.Show(tbr.DocumentObjectType.ToString());
            //            }

            //            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            //            {
            //                DocPicture picture = docObj as DocPicture;
            //                if (picture.AlternativeText == "IS_QRCode")
            //                {
            //                    pictures.Add(docObj);
            //                }
            //            }
            //        }

            //        //if (docObj.DocumentObjectType == DocumentObjectType.TableRow)
            //        //{
            //        //    foreach (DocumentObject docObj2 in docObj.ChildObjects)
            //        //    {

            //        //    }
            //        //}
            //    }

            //    //Replace pitures with the 'newPic'"
            //    foreach (DocumentObject pic in pictures)
            //    {
            //        int index = tb.ChildObjects.IndexOf(pic);
            //        tb.ChildObjects.Insert(index, newPic);
            //        tb.ChildObjects.Remove(pic);
            //    }
            //}

            //doc.SaveToFile("ReplaceWithImage.docx", FileFormat.Docx);

            //Create image list to collect extracted image
            //List<Image> images = new List<Image>();
            //Document doc = new Document();

            //Load a document from disk
            //Image newImage = Image.FromFile(@"E-iceblue.png");
            //DocPicture newPic = new DocPicture(word);
            //newPic.LoadImage(samplePicture);

            //word.LoadFromFile(word);

            //Get the first section
            //Section sec = word.Sections[0];

            ////Get the first table
            //Table table = sec.Tables[1] as Table;
            //for (int i = 0; i < table.Rows.Count; i++)
            //{
            //    for (int j = 0; j < table.Rows[i].Cells.Count; j++)
            //    {
            //        //Foreach paragraph in the cell
            //        foreach (Paragraph par in table.Rows[i].Cells[j].Paragraphs)
            //        {
            //            //Get each document object of paragraph items
            //            foreach (DocumentObject docObject in par.ChildObjects)
            //            {
            //                //If Type of Document Object is Picture, add it into image list
            //                if (docObject.DocumentObjectType == DocumentObjectType.Picture)
            //                {
            //                    DocPicture picture = docObject as DocPicture;
            //                    Image ima = picture.Image;
            //                    images.Add(ima);
            //                }
            //            }
            //        }
            //    }
            //}
        }

        public static void CopyDirectory(string sourceFolder = "", string destFolder = "", string FolderName = "")
        {
            if (!Directory.Exists(destFolder))
            {
                Directory.CreateDirectory(destFolder);
            }
            string[] folders = Directory.GetDirectories(sourceFolder);

            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);
                
                if (name.Contains(FolderName))
                {
                    /*
                      name = name + "_test"; <-- 이것처럼 하면 폴더가 이름 변경되어 진행됨.
                      이곳에서 name을 변경하면 변경된 name으로 폴더가 복사될 것이다.
                      그래서 오과장님이 어떻게 네이밍을 할지에 따라서 수정할지 말지 고민하여 진행한다.
                      변경 안하면 기존 소스처럼, 변경한다고 하시면 그것에 맞게끔 name을 수정하여 개발한다.
                
                      PC에서 자동 스케줄러가 있는지 확인하고 없을 경우 프로그램 수동 실행하는 것으로 말씀드리기.
                     */

                    string dest = Path.Combine(destFolder, name);
                    CopyDirectory(folder, dest);

                    string[] files = Directory.GetFiles(folder);
                    foreach (string file in files)
                    {
                        string name_file = Path.GetFileName(file);
                        string dest_file = Path.Combine(dest, name_file);
                        File.Copy(file, dest_file, true); // 3번째 인자에 true하면 덮어쓰기, 기본은 false
                    }
                }
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                string filepath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\QRFolderSettings.txt";
                string[] lines = File.ReadAllLines(filepath);
                string sCopyPath = "", sPastePath = "", sFolderName = "";                

                foreach (string sPath in lines)
                {
                    if (sPath.Contains("Before Conversion:"))
                    {
                        sCopyPath = sPath.Substring(sPath.IndexOf(':') + 1).Trim();
                    }
                    else if (sPath.Contains("After Conversion:"))
                    {
                        sPastePath = sPath.Substring(sPath.IndexOf(':') + 1).Trim();
                    }
                    else if (sPath.Contains("FolderName1(Contains):"))
                    {
                        sFolderName = sPath.Substring(sPath.IndexOf(':') + 1).Trim();
                    }
                    else
                    {
                        Console.WriteLine("Nothing!");
                    }
                }

                if (!string.IsNullOrEmpty(sCopyPath) && !string.IsNullOrEmpty(sPastePath))
                {
                    string[] files = Directory.GetFiles(sCopyPath);
                    foreach (string file in files)
                    {
                        // word 파일만 변환 그외에는 예외처리
                        if (!file.ToLower().Contains(".docx"))
                        {
                            continue;
                        }

                        string name_file = Path.GetFileName(file);
                        string s_fileNAme_Only = Path.GetFileNameWithoutExtension(name_file);
                        string s_NAme_pdf = Path.GetFileNameWithoutExtension(name_file) + ".pdf";

                        Document document = new Document();
                        document.LoadFromFile(file);
                        replaceQr(document, createQrCode(s_fileNAme_Only), file);

                        string source_file = Path.Combine(sCopyPath, s_NAme_pdf);
                        string dest_file = Path.Combine(sPastePath, s_NAme_pdf);

                        try
                        {
                            // 중복되는 파일이 있어 Move가 불가할 때 Continue 되어 남은 파일들을 변환한다. 만약 덮어쓰기를 원할 경우 추가개발 필요.
                            File.Move(source_file, dest_file);
                        }
                        catch (Exception f)
                        {
                            Console.WriteLine(f.Message);
                            continue;
                        }
                    }
                }

                // MessageBox가 최상위에 노출되도록 한다. 어떤 프로세스를 클릭하던지 상관없다. 계속해서 유지됨.
                MessageBox.Show(new Form() { TopMost = true }, "시험성적서 QR코드 생성 프로그램 실행이 완료 되었습니다.", "알림");

                // 프로그램 종료
                this.Close();
            }
            catch (Exception f)
            {
                MessageBox.Show("Error가 발생하였습니다." + Environment.NewLine + "담당자에게 캡쳐하여 공유 부탁드립니다." + Environment.NewLine + Environment.NewLine + f.Message + Environment.NewLine + Environment.NewLine + f.Source);
                this.Close();
            }
        }
    }
}
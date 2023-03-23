using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Ghostscript.NET.Rasterizer;
using Image = System.Drawing.Image;
using System.Drawing;
using System;
using System.Data.SqlTypes;
using System.Collections;

namespace RPABuildtech2
{
    public class Commands
    {
        public static void RunReport(string pathFolder, string city, bool checkCalc, bool checkPhotos)
        {
            if (Directory.Exists(pathFolder))
            {
                if (checkCalc == true)
                {
                    ConvertPDF2PNG(pathFolder);
                    CreateReportCalc(pathFolder, city, checkCalc);
                }

                if (checkPhotos == true)
                {
                    ConvertPNG2JPG(pathFolder);
                    CreateReportPhotos(pathFolder, city, checkPhotos);
                }
            }
            else
                MessageBox.Show("Folder not Exist");

        }

        public static void CreateReportCalc(string pathFolder, string city, bool checkCalc)
        {
            var pathFolderCalc = Directory.GetDirectories(pathFolder, "CALC*", SearchOption.TopDirectoryOnly)[0];
            string[] filesPath = Directory.GetFiles(pathFolderCalc, "*.png", SearchOption.AllDirectories);
            var pathExe = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            var pathProject = Path.GetDirectoryName(pathExe);
            var pathReportCalc = Path.Combine(pathProject, "MODELO_CALCULOS.docx");

            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Add(pathReportCalc);

            wordApp.Visible = false;

            var filesPathSorted = new SortedDictionary<double, string>();

            foreach (var item in filesPath)
            {
                try
                {
                    var postNumber = float.Parse(Path.GetFileNameWithoutExtension(item), CultureInfo.InvariantCulture.NumberFormat);
                    filesPathSorted.Add(postNumber, item);
                }
                catch
                {
                    var numberString = Path.GetFileNameWithoutExtension(item);
                    var number = numberString.Substring(0, numberString.Length - 1);
                    var postNumber = double.Parse(number, CultureInfo.InvariantCulture.NumberFormat);
                    filesPathSorted.Add(postNumber + 0.5, item);
                }
            }

            foreach (string file in filesPathSorted.Values)
            {
                
                if (File.Exists(file))
                {
                    var postNumber = Path.GetFileNameWithoutExtension(file);
                    var imagePara = doc.Content.Paragraphs.Add();
                    var inlineShape = imagePara.Range.InlineShapes.AddPicture(file);
                    inlineShape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    inlineShape.Width = 700;
                    inlineShape.Height = 450;

                    var labelPara = doc.Content.Paragraphs.Add();
                    labelPara.Range.Text = postNumber;
                    labelPara.Range.Font.Bold = 1;
                    labelPara.Range.Font.Size = 14;
                    labelPara.Format.SpaceBefore = 10;
                    labelPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    if (!filesPathSorted.Values.Last().Equals(file))
                    {
                        labelPara.Range.InsertParagraphAfter();
                        labelPara.Range.InsertBreak(WdBreakType.wdPageBreak);
                    }
                }
            }

            var pathNameReportFinal = Path.Combine(pathFolder, city.ToUpper() + "_"+"CALCULOS");
            doc.SaveAs2(pathNameReportFinal, WdSaveFormat.wdFormatPDF);

            doc.Close(false);
            wordApp.Quit();

            AutoClosingMessageBox.Show("Calculation Report created successfully", timeout: 5);
        }

        public static void CreateReportPhotos(string pathFolder, string city, bool checkCalc){
            
            var pathFolderPhotos = Directory.GetDirectories(pathFolder, "RELA*", SearchOption.TopDirectoryOnly)[0];
            string[] filesPath = Directory.GetFiles(pathFolderPhotos, "*.jpeg", SearchOption.AllDirectories);
            var pathExe = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            var pathProject = Path.GetDirectoryName(pathExe);
            var pathReportPhotos = Path.Combine(pathProject, "RELATORIO FOTOGRAFICO MODELO A4.docx");

            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Add(pathReportPhotos);

            wordApp.Visible = false;

            var filesPathSorted = new SortedDictionary<double, string>();

            foreach (var item in filesPath)
            {
                var postName = Path.GetFileNameWithoutExtension(item).Split('(')[1].Replace(')',' ');
                var postNumber = float.Parse(postName, CultureInfo.InvariantCulture.NumberFormat);
                filesPathSorted.Add(postNumber, item);
            }

            foreach (Range range in doc.StoryRanges)
            {

                Word.Find find = range.Find;
                object findText = "CIDADE";
                object replacText = city.ToUpper();
                object matchCase = true;
                object matchWholeWord = true;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object nmatchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = 2;
                object wrap = WdFindWrap.wdFindContinue;
                object replaceAll = WdReplace.wdReplaceAll;
                find.Execute(
                    ref findText, 
                    ref matchCase, 
                    ref matchWholeWord, 
                    ref matchWildCards, 
                    ref matchSoundsLike,
                    ref nmatchAllWordForms, 
                    ref forward,
                    ref wrap, 
                    ref format, 
                    ref replacText,
                    ref replaceAll, 
                    ref matchKashida,
                    ref matchDiacritics, 
                    ref matchAlefHamza,
                    ref matchControl);
            }

            foreach (var file in filesPathSorted.Values)
            {
                if (File.Exists(file))
                {
                    var postNumber = Path.GetFileNameWithoutExtension(file);
                    var imagePara = doc.Content.Paragraphs.Add();
                    var inlineShape = imagePara.Range.InlineShapes.AddPicture(file);
                    inlineShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    inlineShape.Width = 700;
                    inlineShape.Height = 450;
                }
                //long sizeFile = new FileInfo(pathNameReportFinal + ".pdf").Length;
            }
                
            var pathNameReportFinal = Path.Combine(pathFolder, city.ToUpper() + "_" + "FOTOS");
            doc.SaveAs2(pathNameReportFinal, WdSaveFormat.wdFormatPDF);

            doc.Close(false);
            wordApp.Quit();

            AutoClosingMessageBox.Show("Photos Report created successfully", timeout: 25);

            SplitPDF(pathNameReportFinal + ".pdf", pathFolder, pathNameReportFinal);

        }

        public static void SplitPDF(string pathFile, string pathFolder, string pathNameReportFinal)
        {
            var pathFolderFotos = Path.Combine(pathFolder, "__FOTOS");
            Directory.CreateDirectory(pathFolderFotos);

            if (File.Exists(pathFile)){

                long sizeFilePrincipal = new FileInfo(pathFile).Length;

                if (sizeFilePrincipal > 9700000)
                {

                    var pdfFile = new SautinSoft.PdfMetamorphosis();
                    pdfFile.SplitPDFFileToPDFFolder(pathFile, pathFolderFotos);

                    var pathFolderTemporary = new DirectoryInfo(pathFolderFotos);
                    var files = pathFolderTemporary.GetFiles("*.pdf");
                    var i = 0;
                    long totalSize = 0;
                    var listFiles = new List<string>();

                    foreach (var file in files)
                    {
                        long size = new FileInfo(file.FullName).Length;
                        totalSize += size;
                        listFiles.Add(file.FullName);

                        if (totalSize > 9000000)
                        {
                            var pdfDataList = new ArrayList();

                            foreach (var item in listFiles)
                            {
                                pdfDataList.Add(File.ReadAllBytes(item));
                            }

                            byte[] singlePdfBytes = pdfFile.MergePDFStreamArrayToPDFStream(pdfDataList);

                            i += 1;
                            var namePDF = pathNameReportFinal + "_" + i.ToString() + ".pdf";

                            if (singlePdfBytes != null)
                                File.WriteAllBytes(namePDF, singlePdfBytes);

                            totalSize = 0;
                            listFiles.Clear();
                        }

                    }

                }
            }
            else {
                MessageBox.Show("File not Exist");
            }
            
            if (Directory.Exists(pathFolderFotos))
                Directory.Delete(pathFolderFotos);

            if (File.Exists(pathFile))
                File.Delete(pathFile);

        }
        
        public static void ConvertPDF2PNG(string pathFolder) {

            var pathFolderCalc = Directory.GetDirectories(pathFolder, "CALC*", SearchOption.TopDirectoryOnly)[0];
            string[] filesPath = Directory.GetFiles(pathFolderCalc);

            foreach (var item in filesPath)
            {
                if (!item.ToLower().Contains(".png")) {

                    int dpi = 300;

                    var rasterizer = new GhostscriptRasterizer();

                    var buffer = File.ReadAllBytes(item);
                    var ms = new MemoryStream(buffer);

                    rasterizer.Open(ms);
                    var postNumber = Path.GetFileNameWithoutExtension(item);

                    string pageFilePath = Path.Combine(pathFolderCalc, postNumber + ".png");

                    Image img = rasterizer.GetPage(dpi, 1);
                    img.Save(pageFilePath, ImageFormat.Png);
                    File.Delete(item);
                }
            }
        }

        public static void ConvertPNG2JPG(string pathFolder){

            var pathFolderCalc = Directory.GetDirectories(pathFolder, "RELAT*", SearchOption.TopDirectoryOnly)[0];
            string[] filesPath = Directory.GetFiles(pathFolderCalc);

            foreach (var item in filesPath)
            {
                if (!item.ToLower().Contains(".jpeg"))
                    if(item.ToLower().Contains(".jpg") || item.ToLower().Contains(".png"))
                    {
                        
                        using (Image image = Image.FromFile(item))
                        {
                            var bitmap = new Bitmap(image.Width, image.Height);
                            bitmap.SetResolution(image.HorizontalResolution, image.VerticalResolution);

                            using (var g = Graphics.FromImage(bitmap))
                            {
                                g.Clear(Color.White);
                                g.DrawImageUnscaled(image, 0, 0);
                            }

                            var postNumber = Path.GetFileNameWithoutExtension(item);
                            string pageFilePath = Path.Combine(pathFolderCalc, postNumber + ".jpeg");

                            bitmap.Save(pageFilePath, ImageFormat.Jpeg);
                        }
                        File.Delete(item);
                    }
            }
        }
    }

 }


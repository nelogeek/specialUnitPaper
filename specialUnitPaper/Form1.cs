using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace specialUnitPaper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

           
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            // Получаем текущую дату
            DateTime currentDate = DateTime.Now;

            // Присваиваем отформатированную дату в TextBox
            dateTextBox.Text = currentDate.ToString("dd.MM.yyyy");

            //Application_Startup();
            copyFont();
        }

        private void copyFont()
        {
            // Путь к папке Fonts на диске C:
            string targetFolderPath = @"C:\Fonts";

            // Создаем папку, если ее нет
            if (!Directory.Exists(targetFolderPath))
            {
                Directory.CreateDirectory(targetFolderPath);
            }

            // Путь к файлу в папке проекта
            string sourceFilePath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Times_New_Roman.ttf");

            // Имя файла для копирования
            string fileName = System.IO.Path.GetFileName(sourceFilePath);

            // Полный путь к файлу в папке Fonts на диске C:
            string targetFilePath = System.IO.Path.Combine(targetFolderPath, fileName);

            // Копируем файл, если его еще нет в целевой папке
            if (!File.Exists(targetFilePath))
            {
                File.Copy(sourceFilePath, targetFilePath);
            }
        }


        string selectedFilePath;
        string newFilePath;


        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            infoLabel.Text = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word documents (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;

                    infoLabel.Text = selectedFilePath;

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            infoLabel.Text = "Процесс маркировки запущен";

            //func();
            func_pdf();

            infoLabel.Text = "Генерация завершена";

            //openDoc(newFilePath);

        }

        public void OpenPDFDocument(string pdfFilePath)
        {
            try
            {
                Process.Start(pdfFilePath);
            }
            catch (Exception ex)
            {
                // Обработка ошибок, например, если файл не найден или возникла ошибка при открытии
                Console.WriteLine("Ошибка при открытии PDF: " + ex.Message);
            }
        }

        private void func()
        {

            try
            {


                string sourceFolderPath = System.IO.Path.GetDirectoryName(selectedFilePath);

                string fileName = System.IO.Path.GetFileName(selectedFilePath);

                newFilePath = System.IO.Path.Combine(sourceFolderPath, "копия_" + fileName);

                File.Copy(selectedFilePath, newFilePath, true);

                // Конвертировать документ в PDF
                string pdfFilePath = ConvertDocxToPdf(newFilePath);

                // Добавить пустые страницы через одну
                string modifiedPdfFilePath = AddEmptyPages(pdfFilePath);

                // Конвертировать PDF обратно в документ Word
                string DocxFilePath = ConvertToDocx(modifiedPdfFilePath, newFilePath);

                addFooters(DocxFilePath);

                deleteTempFiles(pdfFilePath, modifiedPdfFilePath);

                Console.WriteLine("Процесс завершен.");
            }
            catch (FileNotFoundException ex)
            {
                // Обработка исключения, когда файл не найден
                MessageBox.Show("Файл не найден: " + ex.FileName, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void func_pdf()
        {

            string sourceFolderPath = System.IO.Path.GetDirectoryName(selectedFilePath);

            string fileName = System.IO.Path.GetFileName(selectedFilePath);

            newFilePath = System.IO.Path.Combine(sourceFolderPath, "копия_" + fileName);

            File.Copy(selectedFilePath, newFilePath, true);

            // Конвертировать документ в PDF
            string pdfFilePath = ConvertDocxToPdf(newFilePath);

            // Добавить пустые страницы через одну
            string modifiedPdfFilePath = AddEmptyPages(pdfFilePath);

            // Добавить колонтитулы
            string finalPath = addFooters_pdf(modifiedPdfFilePath);

            deleteTempFiles_pdf(newFilePath, pdfFilePath, modifiedPdfFilePath);

            OpenPDFDocument(finalPath);

            Console.WriteLine("Процесс завершен.");

        }

        static void deleteTempFiles_pdf(string newFilePath, string pdfFilePath, string modifiedPdfFilePath)
        {
            DeleteFileIfExists(newFilePath);
            DeleteFileIfExists(pdfFilePath);
            DeleteFileIfExists(modifiedPdfFilePath);
        }

        static void deleteTempFiles(string pdfFilePath, string modifiedPdfFilePath)
        {
            DeleteFileIfExists(pdfFilePath);
            DeleteFileIfExists(modifiedPdfFilePath);

        }

        static void DeleteFileIfExists(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine($"Файл {filePath} удален");
                }
                else
                {
                    Console.WriteLine($"Файл {filePath} не удален");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка удаления временных файлов: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private string addFooters_pdf(string path)
        {
            try
            {
                string oldFile = System.IO.Path.GetFullPath(path);
                string watermarkedFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), "done_" + System.IO.Path.GetFileName(path).Replace("modified_", "").Replace("копия_", ""));
                // Creating watermark on a separate layer
                // Creating iTextSharp.text.pdf.PdfReader object to read the Existing PDF Document
                PdfReader reader1 = new PdfReader(oldFile);
                using (FileStream fs = new FileStream(watermarkedFile, FileMode.Create, FileAccess.Write, FileShare.None))
                // Creating iTextSharp.text.pdf.PdfStamper object to write Data from iTextSharp.text.pdf.PdfReader object to FileStream object
                using (PdfStamper stamper = new PdfStamper(reader1, fs))
                {
                    // Getting total number of pages of the Existing Document
                    int pageCount = reader1.NumberOfPages;

                    // Create New Layer for Watermark
                    PdfLayer layer = new PdfLayer("Layer", stamper.Writer);

                    string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    //MessageBox.Show(projectDirectory);
                    //string fontsFolderPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
                    string fontPath = "C:\\Fonts\\Times_New_Roman.ttf" ;
                    BaseFont baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);


                    // Loop through each Page
                    for (int i = 1; i <= pageCount; i++)
                    {
                        // Getting the Page Size
                        iTextSharp.text.Rectangle rect = reader1.GetPageSize(i);

                        // Get the ContentByte object
                        PdfContentByte cb = stamper.GetOverContent(i);

                        // Tell the cb that the next commands should be "bound" to this new layer
                        cb.BeginLayer(layer);

                        BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                        iTextSharp.text.Font font2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL);

                        string fontPath2 = @"C:\Fonts\Times_New_Roman.ttf";
                        BaseFont baseFont2 = BaseFont.CreateFont(fontPath2, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

                        cb.SetColorFill(BaseColor.BLACK);
                        cb.SetFontAndSize(baseFont2, 10);

                        cb.BeginText();

                        float up = 20;
                        float padding = 30;
                        if (checkBox_doublePrint.Checked)
                        {

                            if ((i + StartNumberNumeric.Value) % 2 == 0)
                            {
                                cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, textBox2.Text, 560f - padding, 28f + up, 0);
                                cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, DateTime.TryParse(dateTextBox.Text, out DateTime date) ? date.ToString("dd.MM.yyyy") : "Invalid date", 560f - padding, 15f + up, 0);
                            }
                            else
                            {
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, $"{textBox_footer.Text} /{i + StartNumberNumeric.Value - 1}", 35f + padding, 28f + up, 0);
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, DateTime.TryParse(dateTextBox.Text, out DateTime date) ? date.ToString("dd.MM.yyyy") : "Invalid date", 35f + padding, 15f + up, 0);
                            }

                        }
                        else
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, $"{textBox_footer.Text} /{i + StartNumberNumeric.Value - 1}", 35f + padding, 28f + up, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, DateTime.TryParse(dateTextBox.Text, out DateTime date) ? date.ToString("dd.MM.yyyy") : "Invalid date", 35f + padding, 15f + up, 0);
                        }


                        cb.EndText();

                        // Close the layer
                        cb.EndLayer();
                    }
                }
                reader1.Close();

                return watermarkedFile;
            }

            catch (Exception ex)
            {
                // Обработка исключения, когда файл не найден
                MessageBox.Show("Ошибка добавления колонтитула: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "-1";
            }
        }


        private void addFooters(string filePath)
        {
            Console.WriteLine("Добавление колонтитулов");
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                // Открываем документ
                doc = wordApp.Documents.Open(filePath);
                {
                    /*Visible = true,
                    ScreenUpdating = true*/
                };
                if (checkBox_doublePrint.Checked)
                {
                    addFootersDoublePrint(doc);
                    //TODO сделать независимые колонтитулы
                    //addIndependentFootersToAllPages(doc);
                }
                else
                {
                    addFootersSinglePrint(doc);
                }
                doc.Save();
            }
            catch (Exception ex) { MessageBox.Show("Произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                // Закрываем документ и приложение

                doc?.Close();
                wordApp.Quit();
                ReleaseComObject(doc);
                ReleaseComObject(wordApp);
            }
        }




        private void addFootersDoublePrint(Word.Document doc)
        {
            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

            doc.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  -  настройка: четные-нечетные страницы

            // Футер для нечетных страниц
            Word.HeaderFooter oddPageFooter = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

            oddPageFooter.LinkToPrevious = false;
            oddPageFooter.PageNumbers.RestartNumberingAtSection = true;
            oddPageFooter.PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            // Футер для четных страниц
            Word.HeaderFooter evenPageFooter = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages];

            evenPageFooter.LinkToPrevious = false;
            evenPageFooter.PageNumbers.RestartNumberingAtSection = true;
            evenPageFooter.PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            #region колонтитул нечетной страницы
            // колонтитул нечетной страницы
            Word.Paragraph oddPageFooterParagraph = oddPageFooter.Range.Paragraphs.Add();
            oddPageFooterParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            // Создаем таблицу в футере (1 строка, 2 столбца)
            Word.Table footerTable = oddPageFooter.Range.Tables.Add(oddPageFooterParagraph.Range, 1, 2, ref defaultTableBehavior, ref autoFitBehavior);

            footerTable.Borders.Enable = 0;

            // Добавляем текст и дату в левый столбец
            Word.Range rangeCell = footerTable.Cell(1, 2).Range;

            Word.Paragraph paragraphCell = rangeCell.Paragraphs.Add(rangeCell);

            Word.Range paragraphRange = paragraphCell.Range;

            // Вставляем текст и дату до и после поля номера страницы
            //paragraphRange.Fields.Add(paragraphRange, Word.WdFieldType.wdFieldPage, "page", false);
            rangeCell.InsertBefore($"{textBox2.Text}");
            rangeCell.InsertAfter(/*"\n" + */ dateTextBox.Text.Replace(',', '.'));


            rangeCell.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            // Устанавливаем шрифт и размер для текста в левом столбце
            Word.Font oddPageFooterFont = footerTable.Cell(1, 1).Range.Font;
            oddPageFooterFont.Name = "Arial";
            oddPageFooterFont.Size = 10;
            #endregion



            #region колонтитул четной страницы
            Word.Paragraph evenPageFooterParagraph = evenPageFooter.Range.Paragraphs.Add();
            evenPageFooterParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            // Создаем таблицу в футере (1 строка, 2 столбца)
            Word.Table footerTablen = evenPageFooter.Range.Tables.Add(evenPageFooterParagraph.Range, 1, 2, ref defaultTableBehavior, ref autoFitBehavior);

            footerTablen.Borders.Enable = 0;

            // Добавляем текст и дату в правый столбец
            // Добавляем поле для номера страницы
            Word.Range rangeCelln = footerTablen.Cell(1, 1).Range;

            paragraphCell = rangeCelln.Paragraphs.Add(rangeCelln);

            paragraphRange = paragraphCell.Range;

            // Вставляем текст и дату после поля номера страницы
            //paragraphRange.Fields.Add(paragraphRange, WdFieldType.wdFieldEmpty, "PAGE \\* MERGEFORMAT", true);
            /*Field field = paragraphRange.Fields.Add(paragraphRange, WdFieldType.wdFieldExpression, "{ PAGE }/2");

            field.Update();*/

            Field field = paragraphRange.Fields.Add(paragraphRange, WdFieldType.wdFieldExpression, "PAGE");

            field.Update();

            //rangeCelln.InsertAfter(field.Result.Text);
            //rangeCelln.InsertBefore($"{textBox_footer.Text} /");
            //rangeCelln.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

            rangeCelln.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            // Устанавливаем шрифт и размер для текста в правом столбце
            Word.Font evenPageFooterFont = footerTablen.Cell(1, 1).Range.Font;
            evenPageFooterFont.Name = "Arial";
            evenPageFooterFont.Size = 10;
            #endregion


            #region отступ поля колонтитула
            /*float margin = 20f;*/
            /*oddPageFooter.Range.PageSetup.LeftMargin = margin;
            oddPageFooter.Range.PageSetup.RightMargin = margin;
            evenPageFooter.Range.PageSetup.LeftMargin = margin;
            evenPageFooter.Range.PageSetup.RightMargin = margin;*/

            #endregion
        }



        private void addIndependentFootersToAllPages(Word.Document doc)
        {
            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

            for (int pageIndex = 2; pageIndex <= doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticPages + 1); pageIndex++)
            {
                Word.Range currentPageRange = doc.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, pageIndex);

                // Создаем новую секцию для текущей страницы
                Word.Section currentSection = currentPageRange.Sections[1];

                // Создаем уникальные футеры для каждой секции
                Word.HeaderFooter pageFooter = currentSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                pageFooter.LinkToPrevious = false; // Отключаем связь с предыдущими колонтитулами

                // Очищаем содержимое текущего футера (если нужно)
                pageFooter.Range.Text = "";

                // Добавляем текст и дату в футер
                Word.Paragraph footerParagraph = pageFooter.Range.Paragraphs.Add();
                footerParagraph.Range.Text = $"{textBox_footer.Text} /\n{dateTextBox.Text.Replace(',', '.')} / {pageIndex}";

                // Выравнивание по центру
                footerParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Добавляем номер страницы в центр футера
                Word.Range pageNumberRange = footerParagraph.Range.Paragraphs.Add().Range;
                pageNumberRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                pageNumberRange.Fields.Add(pageNumberRange, Word.WdFieldType.wdFieldPage, "\\* Arabic", false);
            }
        }



        private void addFootersSinglePrint(Word.Document doc)
        {
            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;


            Word.HeaderFooter footer = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            Word.Range footerRange = footer.Range;

            footer.LinkToPrevious = false;
            footer.PageNumbers.RestartNumberingAtSection = true;
            footer.PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            // Добавляем текст и дату в футер
            Word.Paragraph footerParagraph = footer.Range.Paragraphs.Add();

            footer.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            // Добавляем поле для номера страницы
            Word.Field pageNumberField = footerParagraph.Range.Fields.Add(footerParagraph.Range, Word.WdFieldType.wdFieldPage);

            footerParagraph.Range.InsertBefore($"{textBox2.Text} /");

            // Вставляем текст и дату после поля номера страницы
            footerParagraph.Range.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

            // Устанавливаем шрифт и размер для текста
            Word.Font footerFont = footerParagraph.Range.Font;
            footerFont.Name = "Arial";
            footerFont.Size = 10;

        }



        private void openDoc(string filePath)
        {
            Console.WriteLine("Открытие документа");

            // Создаем приложение Microsoft Word
            Word.Application wordApp = new Word.Application();

            try
            {
                // Открываем документ
                Word.Document doc = wordApp.Documents.Open(filePath);

                // Видимость документа
                wordApp.Visible = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при открытии документа: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Освобождаем ресурсы, даже если возникла ошибка
                //if (wordApp != null)
                //{
                //    wordApp.Quit();
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                //}
            }
        }


        static string ConvertDocxToPdf(string docxFilePath)
        {
            Console.WriteLine("Конвертирование из docx в pdf");
            string pdfFilePath = System.IO.Path.ChangeExtension(docxFilePath, ".pdf");

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                // Открываем документ Word
                doc = wordApp.Documents.Open(docxFilePath);

                // Настройка параметров страницы (формат A4)
                doc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                doc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;

                // Сохраняем как PDF
                doc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при конвертации в PDF: {ex.Message}");
            }
            finally
            {
                // Закрываем документ и приложение
                doc?.Close();
                wordApp.Quit();
                ReleaseComObject(doc);
                ReleaseComObject(wordApp);
            }

            return pdfFilePath;
        }

        // Метод для явного освобождения COM-объектов
        static void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при освобождении COM-объекта: {ex.Message}");
            }
        }

        // Функция для добавления пустых страниц через одну
        private static string AddEmptyPages(string pdfFilePath)
        {
            try
            {
                string modifiedPdfFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pdfFilePath), "modified_" + System.IO.Path.GetFileName(pdfFilePath));
                PdfReader reader = new PdfReader(pdfFilePath);
                PdfStamper stamper = new PdfStamper(reader, new FileStream(modifiedPdfFilePath, FileMode.Create));

                int total = reader.NumberOfPages + 1;
                for (int pageNumber = total; pageNumber > 1; pageNumber--)
                {
                    stamper.InsertPage(pageNumber, PageSize.A4);
                }
                stamper.Close();
                reader.Close();

                return modifiedPdfFilePath;
            }

            catch (Exception ex)
            {
                // Обработка исключения, когда файл не найден
                MessageBox.Show("Ошибка добавления пустых страниц: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return "-1";
            }

        }

        // Функция для конвертации PDF обратно в документ Word
        static string ConvertToDocx(string pdfFilePath, string originalDocxFilePath)
        {
            string finalDocxFilePath = originalDocxFilePath; // Переписываем существующий файл
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(pdfFilePath);
            doc.SaveAs2(finalDocxFilePath);
            doc.Close();
            wordApp.Quit();

            return finalDocxFilePath;
        }


    }
}

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

            func();

            infoLabel.Text = "Генерация завершена";

            openDoc(newFilePath);

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
                string finalDocxFilePath = ConvertToDocx(modifiedPdfFilePath, newFilePath);

                addFooters(finalDocxFilePath);

                deleteTempFiles(pdfFilePath, modifiedPdfFilePath);

                Console.WriteLine("Процесс завершен.");
            }
            catch (FileNotFoundException ex)
            {
                // Обработка исключения, когда файл не найден
                MessageBox.Show("Файл не найден: " + ex.FileName, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        static void deleteTempFiles(string pdfFilePath, string modifiedPdfFilePath)
        {
            DeleteFileIfExists(pdfFilePath);
            DeleteFileIfExists(modifiedPdfFilePath);

        }

        static void DeleteFileIfExists(string filePath)
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

            #region колонтитул четной страницы
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



            #region колонтитул нечетной страницы
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
            paragraphRange.Fields.Add(paragraphRange, Word.WdFieldType.wdFieldPage, "page", false);
            rangeCelln.InsertBefore($"{textBox_footer.Text} /");
            rangeCelln.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

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

using Microsoft.Office.Interop.Word;
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


        private void funcWordDocument(string filePath)
        {
            // Создаем приложение Microsoft Word
            Word.Application wordApp = new Word.Application()
            {
                /*Visible = true,
                ScreenUpdating = true*/
            };

            try
            {
                // Открываем документ
                Document doc = wordApp.Documents.Open(filePath);


                AddEmptyPages(doc);

                if (checkBox_doublePrint.Checked)
                {
                    addFootersDoublePrint(doc);
                }
                else
                {
                    addFootersSinglePrint(doc);
                }


                // Сохраняем и закрываем документ
                doc.Save();
                doc.Close();


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
            finally
            {

                // Закрываем приложение Word
                wordApp.Quit();
            }

        }


        private void AddEmptyPages(Document doc)
        {
            // Получаем количество страниц в документе
            int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);

            // Вычисляем общее количество страниц после добавления пустых страниц
            int totalPageCount = pageCount * 2;

            // Добавляем пустую страницу после каждой существующей страницы
            for (int i = 2; i < totalPageCount; i += 2)
            {
                object breakType = Word.WdBreakType.wdPageBreak;
                Word.Range range = doc.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, i);
                range.InsertBreak(ref breakType);
            }
        }



        private void addFootersDoublePrint(Document doc)
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
            Range rangeCell = footerTable.Cell(1, 2).Range;

            Paragraph paragraphCell = rangeCell.Paragraphs.Add(rangeCell);

            Range paragraphRange = paragraphCell.Range;

            // Вставляем текст и дату после поля номера страницы
            paragraphRange.Fields.Add(paragraphRange, Word.WdFieldType.wdFieldPage, "page", false);
            rangeCell.InsertBefore($"{textBox_footer.Text} /");
            rangeCell.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

            rangeCell.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            // Устанавливаем шрифт и размер для текста в левом столбце
            Word.Font oddPageFooterFont = footerTable.Cell(1, 1).Range.Font;
            oddPageFooterFont.Name = "Arial";
            oddPageFooterFont.Size = 10;
            #endregion



            #region колонтитул нечетной страницы
            Word.Paragraph evenPageFooterParagraph = evenPageFooter.Range.Paragraphs.Add();
            evenPageFooterParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            // Создаем таблицу в футере (1 строка, 2 столбца)
            Table footerTablen = evenPageFooter.Range.Tables.Add(evenPageFooterParagraph.Range, 1, 2, ref defaultTableBehavior, ref autoFitBehavior);

            footerTablen.Borders.Enable = 0;

            // Добавляем текст и дату в правый столбец
            // Добавляем поле для номера страницы
            Range rangeCelln = footerTablen.Cell(1, 1).Range;

            paragraphCell = rangeCelln.Paragraphs.Add(rangeCelln);

            paragraphRange = paragraphCell.Range;

            // Вставляем текст и дату после поля номера страницы
            paragraphRange.Fields.Add(paragraphRange, Word.WdFieldType.wdFieldPage, "page", false);
            rangeCelln.InsertBefore($"{textBox_footer.Text} /");
            rangeCelln.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

            rangeCelln.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

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







        private void addFootersSinglePrint(Document doc)
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

            footerParagraph.Range.InsertBefore($"{textBox_footer.Text} /");

            // Вставляем текст и дату после поля номера страницы
            footerParagraph.Range.InsertAfter("\n" + dateTextBox.Text.Replace(',', '.'));

            // Устанавливаем шрифт и размер для текста
            Word.Font footerFont = footerParagraph.Range.Font;
            footerFont.Name = "Arial";
            footerFont.Size = 10;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            infoLabel.Text = "Процесс маркировки запущен";

            string sourceFolderPath = Path.GetDirectoryName(selectedFilePath);

            string fileName = Path.GetFileName(selectedFilePath);

            string newFilePath = Path.Combine(sourceFolderPath, "копия_" + fileName);

            File.Copy(selectedFilePath, newFilePath, true);

            funcWordDocument(newFilePath);

            infoLabel.Text = "Генерация завершена";

            openDoc(newFilePath);

        }

        private void openDoc(string filePath)
        {
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void StartNumberNumeric_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox_footer_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTextBox_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void checkBox_doublePrint_CheckedChanged(object sender, EventArgs e)
        {

        }




        /*private void addFootersDoublePrint(Document doc)
        {
            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;


            doc.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  -  настройка: четные-нечетные страницы
            Word.Range headerRange = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            headerRange.InsertAfter(" " + DateTime.Now.ToString("dd.MM.yyyy"));

            // колонтитул нечетной страницы
            doc.Tables.Add(headerRange, 2, 2, ref defaultTableBehavior, ref autoFitBehavior);

            Word.Table headerTable = headerRange.Tables[1];

            

            headerTable.Cell(2, 1).Merge(headerTable.Cell(2, 2));

            headerTable.Borders.Enable = 1;
            Word.Range rangePageNum = headerTable.Cell(1,2).Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            Word.Range rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            headerTable.Cell(1, 2).Range.Font.Size = 10;

            
            // колонтитул четных страниц
            headerRange = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;

            doc.Tables.Add(headerRange, 2, 2, ref defaultTableBehavior, ref autoFitBehavior);

            headerTable = headerRange.Tables[1];
            headerTable.Columns[1].PreferredWidth = 4f;
            headerTable.Columns[2].PreferredWidth = 4f;
            headerTable.Cell(2, 1).Merge(headerTable.Cell(2, 2));

            headerTable.Borders.Enable = 1;
            rangePageNum = rangePageNum = headerTable.Cell(1, 2).Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            
            headerTable.Cell(1, 1).Range.Font.Size = 10;
        }*/


        /*private void addFootersDoublePrint(Document doc)
        {
            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

            doc.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true - настройка: четные-нечетные страницы
            Word.Range headerRange = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
            doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            // колонтитул нечетной страницы
            // расположение в правом углу
            headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            doc.Tables.Add(headerRange, 2, 2, ref defaultTableBehavior, ref autoFitBehavior);

            Word.Table headerTable = headerRange.Tables[1];

            // Устанавливаем размер таблицы
            *//*headerTable.PreferredWidth = 3f;*//*
            // Устанавливаем выравнивание таблицы в правом углу
            headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;


            headerTable.Cell(2, 1).Merge(headerTable.Cell(2, 2));

            headerTable.Borders.Enable = 1;

            Word.Range rangePageNum = headerTable.Cell(1, 2).Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            Word.Range rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            headerTable.Cell(1, 1).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
            headerTable.Cell(1, 2).Width = 0f;
            headerTable.Cell(2, 1).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;

            headerTable.Cell(1, 1).Range.Text = textBox_footer.Text;
            headerTable.Cell(2, 1).Range.Text = DateTime.Now.ToString("dd.MM.yyyy");

            headerTable.Cell(1, 2).Range.Font.Size = 10;

            

            // колонтитул четных страниц
            headerRange = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;

            doc.Tables.Add(headerRange, 2, 2, ref defaultTableBehavior, ref autoFitBehavior);

            headerTable = headerRange.Tables[1];

            // Устанавливаем размер таблицы
           *//* headerTable.PreferredWidth = 3f;*//*
            // Устанавливаем выравнивание таблицы в правом углу
            headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;


            headerTable.Cell(2, 1).Merge(headerTable.Cell(2, 2));

            headerTable.Borders.Enable = 1;

            rangePageNum = headerTable.Cell(1, 2).Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            headerTable.Cell(1, 1).Width = 200;
            headerTable.Cell(1, 2).Width = 20;
            

            headerTable.Cell(1, 1).Range.Text = textBox_footer.Text;
            headerTable.Cell(2, 1).Range.Text = DateTime.Now.ToString("dd.MM.yyyy");

            headerTable.Cell(1, 2).Range.Font.Size = 10;

        }*/










    }
}

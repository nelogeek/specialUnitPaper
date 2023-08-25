using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;


using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;

using System.Diagnostics;


namespace specialUnitPaper
{
    public partial class Form1 : Form
    {
        private BackgroundWorker worker;
        string selectedFilePath = null;
        string newFilePath = null;


        public Form1()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
        }

        private void InitializeBackgroundWorker()
        {
            worker = new BackgroundWorker();
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            this.Invoke((MethodInvoker)delegate {
                Disable();
            });

            worker.ReportProgress(2, "Процесс маркировки запущен");

            string sourceFolderPath = System.IO.Path.GetDirectoryName(selectedFilePath);

            string fileName = System.IO.Path.GetFileName(selectedFilePath);

            newFilePath = System.IO.Path.Combine(sourceFolderPath, "копия_" + fileName);

            File.Copy(selectedFilePath, newFilePath, true);


            worker.ReportProgress(10, "Конвертация в PDF");

            // Конвертировать документ в PDF
            string pdfFilePath = ConvertDocxToPdf(newFilePath);
            if (pdfFilePath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

            worker.ReportProgress(30, "Добавление пустых страниц");

            // Добавить пустые страницы через одну
            string modifiedPdfFilePath = AddEmptyPages(pdfFilePath);
            if (modifiedPdfFilePath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

            worker.ReportProgress(52, "Добавление колонтитулов");

            // Добавить колонтитулы
            string finalPath = addFooters_pdf(modifiedPdfFilePath);
            if (finalPath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

            worker.ReportProgress(86, "Удаление временных файлов");

            deleteTempFiles_pdf(newFilePath, pdfFilePath, modifiedPdfFilePath);

            worker.ReportProgress(94, "Открытие документа");

            OpenPDFDocument(finalPath);

            
            this.Invoke((MethodInvoker)delegate {
                Enable();
            });

            worker.ReportProgress(100, "Генерация завершена");
        }

        private void Enable()
        {
            textBox2.Enabled = true;
            textBox_footer.Enabled = true;
            StartNumberNumeric.Enabled = true;
            dateTextBox.Enabled = true;
            checkBox_doublePrint.Enabled = true;
            buttonSelectFile.Enabled = true;
        }

        private void Disable()
        {
            textBox2.Enabled = false;
            textBox_footer.Enabled = false;
            StartNumberNumeric.Enabled = false;
            dateTextBox.Enabled = false;
            checkBox_doublePrint.Enabled = false;
            buttonSelectFile.Enabled = false;
            button1.Enabled = false;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                // Обработка ошибок
                infoLabel.Text = "Генерация завершена с ошибкой";
            }
            else if (e.Cancelled)
            {
                // Обработка отмены операции
                infoLabel.Text = "Генерация отменена";
            }
            else
            {
                // Операция успешно завершена
                infoLabel.Text = "Генерация завершена";
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Обновление прогресса
            infoLabel.Text = e.UserState.ToString();
            progressBar.Value = e.ProgressPercentage;
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
            button1.Enabled = !string.IsNullOrEmpty(infoLabel.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            worker.RunWorkerAsync();

            //infoLabel.Text = "Процесс маркировки запущен";

            //func();
            //func_pdf();

            //infoLabel.Text = "Генерация завершена";

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


        private void func_pdf()
        {

            string sourceFolderPath = System.IO.Path.GetDirectoryName(selectedFilePath);

            string fileName = System.IO.Path.GetFileName(selectedFilePath);

            newFilePath = System.IO.Path.Combine(sourceFolderPath, "копия_" + fileName);

            File.Copy(selectedFilePath, newFilePath, true);

            // Конвертировать документ в PDF
            infoLabel.Text = "Конвертация в pdf";
            string pdfFilePath = ConvertDocxToPdf(newFilePath);
            if (pdfFilePath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

            // Добавить пустые страницы через одну
            infoLabel.Text = "Добавление пустых страниц";
            string modifiedPdfFilePath = AddEmptyPages(pdfFilePath);
            if (modifiedPdfFilePath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

            // Добавить колонтитулы
            infoLabel.Text = "Добавление колонтитулов";
            string finalPath = addFooters_pdf(modifiedPdfFilePath);
            if (finalPath == "-1")
            {
                infoLabel.Text = "Генерация завершена с ошибкой";
                return;
            }

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

                            if (i % 2 != 0)
                            {
                                cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, textBox2.Text, 560f - padding, 28f + up, 0);
                                cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, DateTime.TryParse(dateTextBox.Text, out DateTime date) ? date.ToString("dd.MM.yyyy") : "Invalid date", 560f - padding, 15f + up, 0);
                            }
                            else
                            {
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, $"{textBox_footer.Text} /{i/2 + StartNumberNumeric.Value - 1}", 35f + padding, 28f + up, 0);
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
                doc.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatPDF);

                return pdfFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при конвертации в PDF: {ex.Message}");
                return "-1";
            }
            finally
            {
                // Закрываем документ и приложение
                doc?.Close();
                wordApp.Quit();
                ReleaseComObject(doc);
                ReleaseComObject(wordApp);
            }

            
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

    }
}

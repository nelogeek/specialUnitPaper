﻿namespace specialUnitPaper
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonSelectFile = new System.Windows.Forms.Button();
            this.infoLabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.StartNumberNumeric = new System.Windows.Forms.NumericUpDown();
            this.textBox_footer = new System.Windows.Forms.TextBox();
            this.checkBox_doublePrint = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTextBox = new System.Windows.Forms.MaskedTextBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonSelectFile
            // 
            this.buttonSelectFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonSelectFile.Location = new System.Drawing.Point(22, 218);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(145, 33);
            this.buttonSelectFile.TabIndex = 0;
            this.buttonSelectFile.Text = "Выбрать документ";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // infoLabel
            // 
            this.infoLabel.AutoSize = true;
            this.infoLabel.Location = new System.Drawing.Point(12, 284);
            this.infoLabel.MaximumSize = new System.Drawing.Size(279, 13);
            this.infoLabel.MinimumSize = new System.Drawing.Size(35, 13);
            this.infoLabel.Name = "infoLabel";
            this.infoLabel.Size = new System.Drawing.Size(35, 13);
            this.infoLabel.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // StartNumberNumeric
            // 
            this.StartNumberNumeric.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StartNumberNumeric.Location = new System.Drawing.Point(119, 86);
            this.StartNumberNumeric.Name = "StartNumberNumeric";
            this.StartNumberNumeric.Size = new System.Drawing.Size(48, 22);
            this.StartNumberNumeric.TabIndex = 2;
            this.StartNumberNumeric.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.StartNumberNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // textBox_footer
            // 
            this.textBox_footer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox_footer.Location = new System.Drawing.Point(22, 41);
            this.textBox_footer.Multiline = true;
            this.textBox_footer.Name = "textBox_footer";
            this.textBox_footer.Size = new System.Drawing.Size(258, 26);
            this.textBox_footer.TabIndex = 3;
            this.textBox_footer.Text = "Учтенный лист № 19с";
            // 
            // checkBox_doublePrint
            // 
            this.checkBox_doublePrint.AutoSize = true;
            this.checkBox_doublePrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_doublePrint.Location = new System.Drawing.Point(22, 172);
            this.checkBox_doublePrint.Name = "checkBox_doublePrint";
            this.checkBox_doublePrint.Size = new System.Drawing.Size(168, 20);
            this.checkBox_doublePrint.TabIndex = 4;
            this.checkBox_doublePrint.Text = "Двусторонняя печать";
            this.checkBox_doublePrint.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(185, 218);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 33);
            this.button1.TabIndex = 5;
            this.button1.Text = "Выполнить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(19, 88);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Нумерация с ";
            // 
            // dateTextBox
            // 
            this.dateTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTextBox.Location = new System.Drawing.Point(22, 124);
            this.dateTextBox.Mask = "00.00.0000";
            this.dateTextBox.Name = "dateTextBox";
            this.dateTextBox.Size = new System.Drawing.Size(73, 22);
            this.dateTextBox.TabIndex = 7;
            this.dateTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dateTextBox.ValidatingType = typeof(System.DateTime);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Текст колонтитула";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 306);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkBox_doublePrint);
            this.Controls.Add(this.textBox_footer);
            this.Controls.Add(this.StartNumberNumeric);
            this.Controls.Add(this.infoLabel);
            this.Controls.Add(this.buttonSelectFile);
            this.MinimumSize = new System.Drawing.Size(298, 345);
            this.Name = "Form1";
            this.Text = "Маркировка";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonSelectFile;
        private System.Windows.Forms.Label infoLabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.NumericUpDown StartNumberNumeric;
        private System.Windows.Forms.TextBox textBox_footer;
        private System.Windows.Forms.CheckBox checkBox_doublePrint;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.MaskedTextBox dateTextBox;
        private System.Windows.Forms.Label label2;
    }
}

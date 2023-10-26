namespace ToolFormatExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnFormat = new Button();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog2 = new OpenFileDialog();
            textBox1 = new TextBox();
            richTextBox1 = new RichTextBox();
            label1 = new Label();
            label2 = new Label();
            btnFileTemplate = new Button();
            button3 = new Button();
            SuspendLayout();
            // 
            // btnFormat
            // 
            btnFormat.Location = new Point(618, 87);
            btnFormat.Name = "btnFormat";
            btnFormat.Size = new Size(128, 39);
            btnFormat.TabIndex = 0;
            btnFormat.Text = "Format";
            btnFormat.UseVisualStyleBackColor = true;
            btnFormat.Click += btnFormat_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // openFileDialog2
            // 
            openFileDialog2.FileName = "openFileDialog2";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(173, 49);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(189, 23);
            textBox1.TabIndex = 1;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(27, 173);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(453, 265);
            richTextBox1.TabIndex = 2;
            richTextBox1.Text = "";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(27, 52);
            label1.Name = "label1";
            label1.Size = new Size(75, 15);
            label1.TabIndex = 3;
            label1.Text = "File template";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(27, 155);
            label2.Name = "label2";
            label2.Size = new Size(51, 15);
            label2.TabIndex = 3;
            label2.Text = "File xuất";
            // 
            // btnFileTemplate
            // 
            btnFileTemplate.Location = new Point(384, 52);
            btnFileTemplate.Name = "btnFileTemplate";
            btnFileTemplate.Size = new Size(75, 23);
            btnFileTemplate.TabIndex = 4;
            btnFileTemplate.Text = "Mở";
            btnFileTemplate.UseVisualStyleBackColor = true;
            btnFileTemplate.Click += btnFileTemplate_Click;
            // 
            // button3
            // 
            button3.Location = new Point(84, 151);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 4;
            button3.Text = "button2";
            button3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(button3);
            Controls.Add(btnFileTemplate);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(richTextBox1);
            Controls.Add(textBox1);
            Controls.Add(btnFormat);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnFormat;
        private OpenFileDialog openFileDialog1;
        private OpenFileDialog openFileDialog2;
        private TextBox textBox1;
        private RichTextBox richTextBox1;
        private Label label1;
        private Label label2;
        private Button btnFileTemplate;
        private Button button3;
    }
}
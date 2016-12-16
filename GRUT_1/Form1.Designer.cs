namespace GRUT_1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lstViewer = new System.Windows.Forms.ListBox();
            this.btnTryIt = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // lstViewer
            // 
            this.lstViewer.FormattingEnabled = true;
            this.lstViewer.Location = new System.Drawing.Point(40, 23);
            this.lstViewer.Name = "lstViewer";
            this.lstViewer.Size = new System.Drawing.Size(1161, 316);
            this.lstViewer.TabIndex = 0;
            // 
            // btnTryIt
            // 
            this.btnTryIt.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTryIt.Location = new System.Drawing.Point(71, 369);
            this.btnTryIt.Name = "btnTryIt";
            this.btnTryIt.Size = new System.Drawing.Size(132, 70);
            this.btnTryIt.TabIndex = 1;
            this.btnTryIt.Text = "Try It";
            this.btnTryIt.UseVisualStyleBackColor = true;
            this.btnTryIt.Click += new System.EventHandler(this.btnTryIt_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(994, 369);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(125, 60);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel file|*.xls| Text file:|*.txt";
            this.openFileDialog1.InitialDirectory = "D:\\Spring_2017";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.SelectedPath = "D:\\Spring_2017\\TestFiles";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1248, 466);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnTryIt);
            this.Controls.Add(this.lstViewer);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lstViewer;
        private System.Windows.Forms.Button btnTryIt;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}


namespace MAXS.MOFF.Excel_Workshop
{
    using Microsoft.Vbe.Interop;

    partial class MainWindow
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
            this.bgThread = new System.ComponentModel.BackgroundWorker();
            this.TestBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TestBtn
            // 
            this.TestBtn.Location = new System.Drawing.Point(12, 12);
            this.TestBtn.Name = "TestBtn";
            this.TestBtn.Size = new System.Drawing.Size(75, 23);
            this.TestBtn.TabIndex = 0;
            this.TestBtn.Text = "Test Method ";
            this.TestBtn.UseVisualStyleBackColor = true;
            this.TestBtn.Click += new System.EventHandler(this.TestBtn_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.TestBtn);
            this.Name = "MainWindow";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.ComponentModel.BackgroundWorker bgThread;
        private System.Windows.Forms.Button TestBtn;
    }
}


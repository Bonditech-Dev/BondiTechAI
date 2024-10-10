namespace WindowsFormsTest
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
            this.btnConfigureSettings = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConfigureSettings
            // 
            this.btnConfigureSettings.Location = new System.Drawing.Point(89, 99);
            this.btnConfigureSettings.Name = "btnConfigureSettings";
            this.btnConfigureSettings.Size = new System.Drawing.Size(187, 44);
            this.btnConfigureSettings.TabIndex = 0;
            this.btnConfigureSettings.Text = "Configure Settings...";
            this.btnConfigureSettings.UseVisualStyleBackColor = true;
            this.btnConfigureSettings.Click += new System.EventHandler(this.btnConfigureSettings_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 342);
            this.Controls.Add(this.btnConfigureSettings);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnConfigureSettings;
    }
}


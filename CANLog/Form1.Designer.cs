namespace CANLog
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
            this.textBox_Path = new System.Windows.Forms.TextBox();
            this.button_SelectFile = new System.Windows.Forms.Button();
            this.button_Decode = new System.Windows.Forms.Button();
            this.button_Generate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_Path
            // 
            this.textBox_Path.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Path.Location = new System.Drawing.Point(22, 32);
            this.textBox_Path.Multiline = true;
            this.textBox_Path.Name = "textBox_Path";
            this.textBox_Path.Size = new System.Drawing.Size(337, 27);
            this.textBox_Path.TabIndex = 0;
            // 
            // button_SelectFile
            // 
            this.button_SelectFile.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_SelectFile.Location = new System.Drawing.Point(378, 32);
            this.button_SelectFile.Name = "button_SelectFile";
            this.button_SelectFile.Size = new System.Drawing.Size(104, 27);
            this.button_SelectFile.TabIndex = 1;
            this.button_SelectFile.Text = "Select File";
            this.button_SelectFile.UseVisualStyleBackColor = true;
            this.button_SelectFile.Click += new System.EventHandler(this.button_SelectFileClick);
            // 
            // button_Decode
            // 
            this.button_Decode.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Decode.Location = new System.Drawing.Point(22, 116);
            this.button_Decode.Name = "button_Decode";
            this.button_Decode.Size = new System.Drawing.Size(144, 27);
            this.button_Decode.TabIndex = 2;
            this.button_Decode.Text = "Decode CAN log";
            this.button_Decode.UseVisualStyleBackColor = true;
            this.button_Decode.Click += new System.EventHandler(this.button_Decode_Click);
            // 
            // button_Generate
            // 
            this.button_Generate.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Generate.Location = new System.Drawing.Point(253, 116);
            this.button_Generate.Name = "button_Generate";
            this.button_Generate.Size = new System.Drawing.Size(106, 27);
            this.button_Generate.TabIndex = 3;
            this.button_Generate.Text = "Generate";
            this.button_Generate.UseVisualStyleBackColor = true;
            this.button_Generate.Click += new System.EventHandler(this.button_Generate_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(502, 261);
            this.Controls.Add(this.button_Generate);
            this.Controls.Add(this.button_Decode);
            this.Controls.Add(this.button_SelectFile);
            this.Controls.Add(this.textBox_Path);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_Path;
        private System.Windows.Forms.Button button_SelectFile;
        private System.Windows.Forms.Button button_Decode;
        private System.Windows.Forms.Button button_Generate;
    }
}


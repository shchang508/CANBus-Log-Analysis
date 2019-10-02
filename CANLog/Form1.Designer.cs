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
            this.textBox_Path = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.button_SelectFile = new MaterialSkin.Controls.MaterialFlatButton();
            this.button_Decode = new System.Windows.Forms.Button();
            this.button_Generate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_Path
            // 
            this.textBox_Path.BackColor = System.Drawing.SystemColors.Control;
            this.textBox_Path.Depth = 0;
            this.textBox_Path.Enabled = false;
            this.textBox_Path.Font = new System.Drawing.Font("Microsoft JhengHei", 12F);
            this.textBox_Path.Hint = "";
            this.textBox_Path.Location = new System.Drawing.Point(43, 113);
            this.textBox_Path.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBox_Path.MaxLength = 32767;
            this.textBox_Path.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox_Path.Name = "textBox_Path";
            this.textBox_Path.PasswordChar = '\0';
            this.textBox_Path.SelectedText = "";
            this.textBox_Path.SelectionLength = 0;
            this.textBox_Path.SelectionStart = 0;
            this.textBox_Path.Size = new System.Drawing.Size(356, 23);
            this.textBox_Path.TabIndex = 4;
            this.textBox_Path.TabStop = false;
            this.textBox_Path.UseSystemPasswordChar = false;
            // 
            // button_SelectFile
            // 
            this.button_SelectFile.AutoSize = true;
            this.button_SelectFile.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_SelectFile.Depth = 0;
            this.button_SelectFile.Icon = null;
            this.button_SelectFile.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button_SelectFile.Location = new System.Drawing.Point(399, 100);
            this.button_SelectFile.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.button_SelectFile.MouseState = MaterialSkin.MouseState.HOVER;
            this.button_SelectFile.Name = "button_SelectFile";
            this.button_SelectFile.Primary = false;
            this.button_SelectFile.Size = new System.Drawing.Size(100, 36);
            this.button_SelectFile.TabIndex = 7;
            this.button_SelectFile.Text = "Select File";
            this.button_SelectFile.UseVisualStyleBackColor = true;
            this.button_SelectFile.Click += new System.EventHandler(this.button_SelectFileClick);
            // 
            // button_Decode
            // 
            this.button_Decode.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Decode.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Decode.ForeColor = System.Drawing.Color.White;
            this.button_Decode.Location = new System.Drawing.Point(43, 185);
            this.button_Decode.Name = "button_Decode";
            this.button_Decode.Size = new System.Drawing.Size(144, 35);
            this.button_Decode.TabIndex = 8;
            this.button_Decode.Text = "DECODE LOG";
            this.button_Decode.UseVisualStyleBackColor = true;
            this.button_Decode.Click += new System.EventHandler(this.button_Decode_Click);
            // 
            // button_Generate
            // 
            this.button_Generate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Generate.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_Generate.ForeColor = System.Drawing.Color.White;
            this.button_Generate.Location = new System.Drawing.Point(255, 185);
            this.button_Generate.Name = "button_Generate";
            this.button_Generate.Size = new System.Drawing.Size(144, 35);
            this.button_Generate.TabIndex = 9;
            this.button_Generate.Text = "GENERATE REPORT";
            this.button_Generate.UseVisualStyleBackColor = true;
            this.button_Generate.Click += new System.EventHandler(this.button_Generate_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(547, 267);
            this.Controls.Add(this.button_Generate);
            this.Controls.Add(this.button_Decode);
            this.Controls.Add(this.button_SelectFile);
            this.Controls.Add(this.textBox_Path);
            this.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Sizable = false;
            this.Text = "Report Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox_Path;
        private MaterialSkin.Controls.MaterialFlatButton button_SelectFile;
        private System.Windows.Forms.Button button_Decode;
        private System.Windows.Forms.Button button_Generate;
    }
}


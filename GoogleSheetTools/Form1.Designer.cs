
namespace GoogleSheetTools
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Convert = new System.Windows.Forms.Button();
            this.urlText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.ConvertAll = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Convert
            // 
            this.Convert.Location = new System.Drawing.Point(260, 181);
            this.Convert.Name = "Convert";
            this.Convert.Size = new System.Drawing.Size(75, 23);
            this.Convert.TabIndex = 0;
            this.Convert.Text = "Convert";
            this.Convert.UseVisualStyleBackColor = true;
            this.Convert.Click += new System.EventHandler(this.OnClickConvert);
            // 
            // urlText
            // 
            this.urlText.Location = new System.Drawing.Point(109, 13);
            this.urlText.Name = "urlText";
            this.urlText.Size = new System.Drawing.Size(307, 22);
            this.urlText.TabIndex = 1;
            this.urlText.Text = "1zQuwPyj1tkBBWFGQ464pGvmzjaA1PMIuiQZOZv5WyHg";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "Google Sheet URL";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(335, 41);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(81, 23);
            this.button2.TabIndex = 4;
            this.button2.Text = "Search";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.OnClickSearch);
            // 
            // comboBox
            // 
            this.comboBox.FormattingEnabled = true;
            this.comboBox.Location = new System.Drawing.Point(12, 70);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(404, 20);
            this.comboBox.TabIndex = 5;
            // 
            // ConvertAll
            // 
            this.ConvertAll.Location = new System.Drawing.Point(341, 181);
            this.ConvertAll.Name = "ConvertAll";
            this.ConvertAll.Size = new System.Drawing.Size(75, 23);
            this.ConvertAll.TabIndex = 8;
            this.ConvertAll.Text = "ConvertAll";
            this.ConvertAll.UseVisualStyleBackColor = true;
            this.ConvertAll.Click += new System.EventHandler(this.ConvertAll_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(428, 213);
            this.Controls.Add(this.ConvertAll);
            this.Controls.Add(this.comboBox);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.urlText);
            this.Controls.Add(this.Convert);
            this.Name = "MainForm";
            this.Text = "Google Sheet Tools";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.TextBox urlText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Button ConvertAll;
    }
}


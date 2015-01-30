namespace BalanceSheetUtility
{
    partial class BalanceSheetUtility
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
            this.button1 = new System.Windows.Forms.Button();
            this.LogBox = new System.Windows.Forms.RichTextBox();
            this.Prgresslbl = new System.Windows.Forms.Label();
            this.logclear = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Times New Roman", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(24, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(420, 45);
            this.button1.TabIndex = 0;
            this.button1.Text = "Populate Balance Sheet";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // LogBox
            // 
            this.LogBox.Location = new System.Drawing.Point(24, 112);
            this.LogBox.Name = "LogBox";
            this.LogBox.ReadOnly = true;
            this.LogBox.Size = new System.Drawing.Size(420, 204);
            this.LogBox.TabIndex = 1;
            this.LogBox.Text = "";
            // 
            // Prgresslbl
            // 
            this.Prgresslbl.AutoSize = true;
            this.Prgresslbl.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Prgresslbl.Location = new System.Drawing.Point(21, 92);
            this.Prgresslbl.Name = "Prgresslbl";
            this.Prgresslbl.Size = new System.Drawing.Size(97, 16);
            this.Prgresslbl.TabIndex = 2;
            this.Prgresslbl.Text = "Progress Status";
            this.Prgresslbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // logclear
            // 
            this.logclear.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logclear.Location = new System.Drawing.Point(369, 88);
            this.logclear.Name = "logclear";
            this.logclear.Size = new System.Drawing.Size(75, 23);
            this.logclear.TabIndex = 3;
            this.logclear.Text = "Clear Log";
            this.logclear.UseVisualStyleBackColor = true;
            this.logclear.Click += new System.EventHandler(this.logclear_Click);
            // 
            // BalanceSheetUtility
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(482, 339);
            this.Controls.Add(this.logclear);
            this.Controls.Add(this.Prgresslbl);
            this.Controls.Add(this.LogBox);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BalanceSheetUtility";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BalanceSheetUtility";
            this.Load += new System.EventHandler(this.BalanceSheetUtility_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RichTextBox LogBox;
        private System.Windows.Forms.Label Prgresslbl;
        private System.Windows.Forms.Button logclear;
    }
}


namespace Tara_app
{
    partial class intex_data
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(intex_data));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.deals = new System.Windows.Forms.ComboBox();
            this.yymm = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // deals
            // 
            this.deals.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.deals.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.deals.FormattingEnabled = true;
            this.deals.Location = new System.Drawing.Point(12, 18);
            this.deals.MaxDropDownItems = 20;
            this.deals.Name = "deals";
            this.deals.Size = new System.Drawing.Size(121, 21);
            this.deals.TabIndex = 0;
            this.deals.Text = "transaction";
            this.deals.SelectedIndexChanged += new System.EventHandler(this.deals_SelectedIndexChanged);
            // 
            // yymm
            // 
            this.yymm.FormattingEnabled = true;
            this.yymm.Location = new System.Drawing.Point(12, 45);
            this.yymm.Name = "yymm";
            this.yymm.Size = new System.Drawing.Size(121, 21);
            this.yymm.TabIndex = 1;
            this.yymm.Text = "yymm";
            this.yymm.SelectedIndexChanged += new System.EventHandler(this.yymm_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(158, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "load";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // intex_data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(232, 83);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.yymm);
            this.Controls.Add(this.deals);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(30, 39);
            this.Name = "intex_data";
            this.ShowIcon = false;
            this.Text = "Intex Data";
            this.Load += new System.EventHandler(this.intex_data_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ComboBox deals;
        private System.Windows.Forms.ComboBox yymm;
        private System.Windows.Forms.Button button1;
    }
}
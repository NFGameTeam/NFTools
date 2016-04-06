namespace FileProcess
{
    partial class frmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.skinButton1 = new CCWin.SkinControl.SkinButton();
            this.skinPictureBox1 = new CCWin.SkinControl.SkinPictureBox();
            this.skinPictureBox2 = new CCWin.SkinControl.SkinPictureBox();
            this.skinButton2 = new CCWin.SkinControl.SkinButton();
            this.skinProgressBar1 = new CCWin.SkinControl.SkinProgressBar();
            this.skinProgressBar2 = new CCWin.SkinControl.SkinProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.skinPictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.skinPictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // skinButton1
            // 
            this.skinButton1.BackColor = System.Drawing.Color.Transparent;
            this.skinButton1.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.skinButton1.DownBack = null;
            this.skinButton1.Enabled = false;
            this.skinButton1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.skinButton1.Location = new System.Drawing.Point(157, 176);
            this.skinButton1.MouseBack = null;
            this.skinButton1.Name = "skinButton1";
            this.skinButton1.NormlBack = null;
            this.skinButton1.Size = new System.Drawing.Size(148, 64);
            this.skinButton1.TabIndex = 0;
            this.skinButton1.Text = "XML => MySQL";
            this.skinButton1.UseVisualStyleBackColor = false;
            this.skinButton1.Click += new System.EventHandler(this.OnCreateMysqlFile);
            // 
            // skinPictureBox1
            // 
            this.skinPictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.skinPictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("skinPictureBox1.Image")));
            this.skinPictureBox1.Location = new System.Drawing.Point(74, 176);
            this.skinPictureBox1.Name = "skinPictureBox1";
            this.skinPictureBox1.Size = new System.Drawing.Size(64, 64);
            this.skinPictureBox1.TabIndex = 1;
            this.skinPictureBox1.TabStop = false;
            // 
            // skinPictureBox2
            // 
            this.skinPictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.skinPictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("skinPictureBox2.Image")));
            this.skinPictureBox2.Location = new System.Drawing.Point(74, 79);
            this.skinPictureBox2.Name = "skinPictureBox2";
            this.skinPictureBox2.Size = new System.Drawing.Size(64, 64);
            this.skinPictureBox2.TabIndex = 2;
            this.skinPictureBox2.TabStop = false;
            // 
            // skinButton2
            // 
            this.skinButton2.BackColor = System.Drawing.Color.Transparent;
            this.skinButton2.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.skinButton2.DownBack = null;
            this.skinButton2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.skinButton2.Location = new System.Drawing.Point(157, 81);
            this.skinButton2.MouseBack = null;
            this.skinButton2.Name = "skinButton2";
            this.skinButton2.NormlBack = null;
            this.skinButton2.Size = new System.Drawing.Size(148, 62);
            this.skinButton2.TabIndex = 3;
            this.skinButton2.Text = "Excel => XML";
            this.skinButton2.UseVisualStyleBackColor = false;
            this.skinButton2.Click += new System.EventHandler(this.OnCreateXMLFile);
            // 
            // skinProgressBar1
            // 
            this.skinProgressBar1.Back = null;
            this.skinProgressBar1.BackColor = System.Drawing.Color.Transparent;
            this.skinProgressBar1.BarBack = null;
            this.skinProgressBar1.BarRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinProgressBar1.ForeColor = System.Drawing.Color.Red;
            this.skinProgressBar1.Location = new System.Drawing.Point(7, 264);
            this.skinProgressBar1.Margin = new System.Windows.Forms.Padding(0);
            this.skinProgressBar1.Name = "skinProgressBar1";
            this.skinProgressBar1.Radius = 1;
            this.skinProgressBar1.RadiusStyle = CCWin.SkinClass.RoundStyle.None;
            this.skinProgressBar1.Size = new System.Drawing.Size(346, 13);
            this.skinProgressBar1.TabIndex = 4;
            // 
            // skinProgressBar2
            // 
            this.skinProgressBar2.Back = null;
            this.skinProgressBar2.BackColor = System.Drawing.Color.Transparent;
            this.skinProgressBar2.BarBack = null;
            this.skinProgressBar2.BarRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinProgressBar2.ForeColor = System.Drawing.Color.Red;
            this.skinProgressBar2.Location = new System.Drawing.Point(7, 279);
            this.skinProgressBar2.Name = "skinProgressBar2";
            this.skinProgressBar2.Radius = 1;
            this.skinProgressBar2.RadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinProgressBar2.Size = new System.Drawing.Size(346, 13);
            this.skinProgressBar2.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(139)))), ((int)(((byte)(143)))), ((int)(((byte)(194)))));
            this.ClientSize = new System.Drawing.Size(360, 300);
            this.Controls.Add(this.skinProgressBar2);
            this.Controls.Add(this.skinProgressBar1);
            this.Controls.Add(this.skinButton2);
            this.Controls.Add(this.skinPictureBox2);
            this.Controls.Add(this.skinPictureBox1);
            this.Controls.Add(this.skinButton1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowBorder = false;
            this.ShowDrawIcon = false;
            this.ShowIcon = false;
            this.Text = "NFrame Tool";
            ((System.ComponentModel.ISupportInitialize)(this.skinPictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.skinPictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CCWin.SkinControl.SkinButton skinButton1;
        private CCWin.SkinControl.SkinPictureBox skinPictureBox1;
        private CCWin.SkinControl.SkinPictureBox skinPictureBox2;
        private CCWin.SkinControl.SkinButton skinButton2;
        private CCWin.SkinControl.SkinProgressBar skinProgressBar1;
        private CCWin.SkinControl.SkinProgressBar skinProgressBar2;
    }
}



namespace FindWordsInExcel
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSrcFind = new System.Windows.Forms.Button();
            this.textSrc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textKeyword = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.textResult = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textColum = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProcess = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnSrcFind
            // 
            this.btnSrcFind.Location = new System.Drawing.Point(884, 46);
            this.btnSrcFind.Name = "btnSrcFind";
            this.btnSrcFind.Size = new System.Drawing.Size(97, 40);
            this.btnSrcFind.TabIndex = 0;
            this.btnSrcFind.Text = "Find";
            this.btnSrcFind.UseVisualStyleBackColor = true;
            this.btnSrcFind.Click += new System.EventHandler(this.btnSrcFind_Click);
            // 
            // textSrc
            // 
            this.textSrc.Location = new System.Drawing.Point(191, 46);
            this.textSrc.Name = "textSrc";
            this.textSrc.ReadOnly = true;
            this.textSrc.Size = new System.Drawing.Size(673, 26);
            this.textSrc.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(74, 111);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "Keyword";
            // 
            // textKeyword
            // 
            this.textKeyword.Location = new System.Drawing.Point(191, 100);
            this.textKeyword.Name = "textKeyword";
            this.textKeyword.Size = new System.Drawing.Size(673, 26);
            this.textKeyword.TabIndex = 1;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(884, 100);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(97, 40);
            this.btnRun.TabIndex = 0;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // textResult
            // 
            this.textResult.Location = new System.Drawing.Point(77, 218);
            this.textResult.Multiline = true;
            this.textResult.Name = "textResult";
            this.textResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textResult.Size = new System.Drawing.Size(1092, 391);
            this.textResult.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(74, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 18);
            this.label2.TabIndex = 2;
            this.label2.Text = "Source Folder";
            // 
            // textColum
            // 
            this.textColum.Location = new System.Drawing.Point(191, 153);
            this.textColum.Name = "textColum";
            this.textColum.Size = new System.Drawing.Size(104, 26);
            this.textColum.TabIndex = 1;
            this.textColum.Text = "3";
            this.textColum.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(74, 158);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 18);
            this.label3.TabIndex = 2;
            this.label3.Text = "Target Colum";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(78, 193);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1000, 13);
            this.progressBar1.TabIndex = 4;
            // 
            // lblProcess
            // 
            this.lblProcess.AutoSize = true;
            this.lblProcess.Location = new System.Drawing.Point(1084, 188);
            this.lblProcess.Name = "lblProcess";
            this.lblProcess.Size = new System.Drawing.Size(0, 18);
            this.lblProcess.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1197, 632);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.textResult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblProcess);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textColum);
            this.Controls.Add(this.textKeyword);
            this.Controls.Add(this.textSrc);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnSrcFind);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.Text = "Find words in excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSrcFind;
        private System.Windows.Forms.TextBox textSrc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textKeyword;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox textResult;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textColum;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProcess;
    }
}


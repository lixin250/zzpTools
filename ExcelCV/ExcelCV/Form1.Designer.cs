using System;
using System.Diagnostics;

namespace ExcelCV
{
    partial class 钟志平Ver2
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ButtonA = new System.Windows.Forms.Button();
            this.buttonB = new System.Windows.Forms.Button();
            this.buttonC = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ButtonA
            // 
            this.ButtonA.Location = new System.Drawing.Point(12, 12);
            this.ButtonA.Name = "ButtonA";
            this.ButtonA.Size = new System.Drawing.Size(95, 39);
            this.ButtonA.TabIndex = 0;
            this.ButtonA.Text = "种草1";
            this.ButtonA.UseVisualStyleBackColor = true;
            this.ButtonA.Click += new System.EventHandler(this.ButtonAClick);
            // 
            // buttonB
            // 
            this.buttonB.Location = new System.Drawing.Point(113, 12);
            this.buttonB.Name = "buttonB";
            this.buttonB.Size = new System.Drawing.Size(95, 39);
            this.buttonB.TabIndex = 1;
            this.buttonB.Text = "实时表2";
            this.buttonB.UseVisualStyleBackColor = true;
            this.buttonB.Click += new System.EventHandler(this.ButtonBClick);
            // 
            // buttonC
            // 
            this.buttonC.Location = new System.Drawing.Point(214, 12);
            this.buttonC.Name = "buttonC";
            this.buttonC.Size = new System.Drawing.Size(95, 39);
            this.buttonC.TabIndex = 2;
            this.buttonC.Text = "总表3";
            this.buttonC.UseVisualStyleBackColor = true;
            this.buttonC.Click += new System.EventHandler(this.ButtonCClick);
            // 
            // 钟志平Ver2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 131);
            this.Controls.Add(this.buttonC);
            this.Controls.Add(this.buttonB);
            this.Controls.Add(this.ButtonA);
            this.Name = "钟志平Ver2";
            this.Text = "钟志平Ver2";
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Button btnSelectFiles;
        #endregion

        private System.Windows.Forms.Button ButtonA;
        private System.Windows.Forms.Button buttonB;
        private System.Windows.Forms.Button buttonC;
    }
}


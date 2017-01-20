using System;

namespace PPT_Section_Indicator
{
    public partial class ProgressDialogBox
    {
        const string PROGRESS_MESSAGE = "Processing slide ";
        
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
            this.ProgressRingPictureBox = new System.Windows.Forms.PictureBox();
            this.ProgressMainMessageLabel = new System.Windows.Forms.Label();
            this.ProgressSecondaryMessageLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ProgressRingPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // ProgressRingPictureBox
            // 
            this.ProgressRingPictureBox.Image = global::PPT_Section_Indicator.Properties.Resources.progress_ring;
            this.ProgressRingPictureBox.Location = new System.Drawing.Point(13, 13);
            this.ProgressRingPictureBox.Name = "ProgressRingPictureBox";
            this.ProgressRingPictureBox.Size = new System.Drawing.Size(50, 50);
            this.ProgressRingPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.ProgressRingPictureBox.TabIndex = 0;
            this.ProgressRingPictureBox.TabStop = false;
            // 
            // ProgressMainMessageLabel
            // 
            this.ProgressMainMessageLabel.AutoSize = true;
            this.ProgressMainMessageLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProgressMainMessageLabel.Location = new System.Drawing.Point(83, 13);
            this.ProgressMainMessageLabel.Name = "ProgressMainMessageLabel";
            this.ProgressMainMessageLabel.Size = new System.Drawing.Size(95, 17);
            this.ProgressMainMessageLabel.TabIndex = 3;
            this.ProgressMainMessageLabel.Text = "Please, wait...";
            // 
            // ProgressSecondaryMessageLabel
            // 
            this.ProgressSecondaryMessageLabel.AutoSize = true;
            this.ProgressSecondaryMessageLabel.Location = new System.Drawing.Point(83, 40);
            this.ProgressSecondaryMessageLabel.Name = "ProgressSecondaryMessageLabel";
            this.ProgressSecondaryMessageLabel.Size = new System.Drawing.Size(83, 13);
            this.ProgressSecondaryMessageLabel.TabIndex = 4;
            this.ProgressSecondaryMessageLabel.Text = "Processing slide";
            // 
            // ProgressDialogBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 81);
            this.ControlBox = false;
            this.Controls.Add(this.ProgressSecondaryMessageLabel);
            this.Controls.Add(this.ProgressMainMessageLabel);
            this.Controls.Add(this.ProgressRingPictureBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressDialogBox";
            this.Text = "PPT Section Indicator";
            this.Shown += new System.EventHandler(this.ProgressDialogBox_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.ProgressRingPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox ProgressRingPictureBox;
        private System.Windows.Forms.Label ProgressMainMessageLabel;
        private System.Windows.Forms.Label ProgressSecondaryMessageLabel;

        
    }
}
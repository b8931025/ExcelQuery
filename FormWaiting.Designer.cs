namespace ExcelQuery
{
    partial class FormWaiting
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
            // 
            // lblWating
            // 
            this.lblWating = new System.Windows.Forms.Label();
            this.lblWating.Font = new System.Drawing.Font("Microsoft JhengHei UI", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.lblWating.AutoSize = true;
            this.lblWating.Name = "lblWating";
            this.lblWating.Size = new System.Drawing.Size(67, 15);
            this.lblWating.Text = "Waiting";
            this.lblWating.Location = new System.Drawing.Point((this.ClientSize.Width - lblWating.Size.Width) / 2,
                                                       (this.ClientSize.Height - lblWating.Size.Height) / 2);

            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(lblWating);
            this.Text = "Waiting a moment";
        }

        #endregion
        private Label lblWating;
    }
}
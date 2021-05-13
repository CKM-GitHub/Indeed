
namespace Indeed
{
    partial class Menu
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
            this.btnIndeed = new System.Windows.Forms.Button();
            this.btnBaseConnect = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnIndeed
            // 
            this.btnIndeed.Location = new System.Drawing.Point(37, 21);
            this.btnIndeed.Name = "btnIndeed";
            this.btnIndeed.Size = new System.Drawing.Size(181, 52);
            this.btnIndeed.TabIndex = 0;
            this.btnIndeed.Tag = "1";
            this.btnIndeed.Text = "Indeed";
            this.btnIndeed.UseVisualStyleBackColor = true;
            this.btnIndeed.Click += new System.EventHandler(this.btnIndeed_Click);
            // 
            // btnBaseConnect
            // 
            this.btnBaseConnect.Location = new System.Drawing.Point(37, 107);
            this.btnBaseConnect.Name = "btnBaseConnect";
            this.btnBaseConnect.Size = new System.Drawing.Size(181, 52);
            this.btnBaseConnect.TabIndex = 1;
            this.btnBaseConnect.Tag = "2";
            this.btnBaseConnect.Text = "Base Connect";
            this.btnBaseConnect.UseVisualStyleBackColor = true;
            this.btnBaseConnect.Click += new System.EventHandler(this.btnIndeed_Click);
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(828, 486);
            this.Controls.Add(this.btnBaseConnect);
            this.Controls.Add(this.btnIndeed);
            this.Name = "Menu";
            this.Text = "Menu";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnIndeed;
        private System.Windows.Forms.Button btnBaseConnect;
    }
}
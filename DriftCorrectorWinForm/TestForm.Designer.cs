namespace AsposeFormAdjustment
{
    partial class TestForm
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
            btnConvertToV23PDF = new System.Windows.Forms.Button();
            btnDiffNCompare = new System.Windows.Forms.Button();
            btnOnlyCompare = new System.Windows.Forms.Button();
            btnPolicyDocToV23PDF = new System.Windows.Forms.Button();
            btnPolicyDocToV14PDF = new System.Windows.Forms.Button();
            btnRectifyPolicyDocs = new System.Windows.Forms.Button();
            cmbCustomer = new System.Windows.Forms.ComboBox();
            btnCopyToDestination = new System.Windows.Forms.Button();
            label1 = new System.Windows.Forms.Label();
            txtPolicyTemplateList = new System.Windows.Forms.TextBox();
            label2 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            btnCopyPolicyTemplates = new System.Windows.Forms.Button();
            SuspendLayout();
            // 
            // btnConvertToV23PDF
            // 
            btnConvertToV23PDF.Location = new System.Drawing.Point(38, 20);
            btnConvertToV23PDF.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnConvertToV23PDF.Name = "btnConvertToV23PDF";
            btnConvertToV23PDF.Size = new System.Drawing.Size(185, 22);
            btnConvertToV23PDF.TabIndex = 0;
            btnConvertToV23PDF.Text = "Convert Doc to V23 PDF";
            btnConvertToV23PDF.UseVisualStyleBackColor = true;
            btnConvertToV23PDF.Click += btnConvertToV23PDF_Click;
            // 
            // btnDiffNCompare
            // 
            btnDiffNCompare.Location = new System.Drawing.Point(277, 20);
            btnDiffNCompare.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnDiffNCompare.Name = "btnDiffNCompare";
            btnDiffNCompare.Size = new System.Drawing.Size(185, 22);
            btnDiffNCompare.TabIndex = 1;
            btnDiffNCompare.Text = "DiffNCompare";
            btnDiffNCompare.UseVisualStyleBackColor = true;
            btnDiffNCompare.Click += btnDiffNCompare_Click;
            // 
            // btnOnlyCompare
            // 
            btnOnlyCompare.Location = new System.Drawing.Point(488, 20);
            btnOnlyCompare.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnOnlyCompare.Name = "btnOnlyCompare";
            btnOnlyCompare.Size = new System.Drawing.Size(185, 22);
            btnOnlyCompare.TabIndex = 2;
            btnOnlyCompare.Text = "Only Compare";
            btnOnlyCompare.UseVisualStyleBackColor = true;
            btnOnlyCompare.Click += btnOnlyCompare_Click;
            // 
            // btnPolicyDocToV23PDF
            // 
            btnPolicyDocToV23PDF.Location = new System.Drawing.Point(313, 84);
            btnPolicyDocToV23PDF.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnPolicyDocToV23PDF.Name = "btnPolicyDocToV23PDF";
            btnPolicyDocToV23PDF.Size = new System.Drawing.Size(113, 43);
            btnPolicyDocToV23PDF.TabIndex = 3;
            btnPolicyDocToV23PDF.Text = "Convert Policy Doc to V23 PDF";
            btnPolicyDocToV23PDF.UseCompatibleTextRendering = true;
            btnPolicyDocToV23PDF.UseVisualStyleBackColor = true;
            btnPolicyDocToV23PDF.Click += btnPolicyDocToV23PDF_Click;
            // 
            // btnPolicyDocToV14PDF
            // 
            btnPolicyDocToV14PDF.Location = new System.Drawing.Point(193, 84);
            btnPolicyDocToV14PDF.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnPolicyDocToV14PDF.Name = "btnPolicyDocToV14PDF";
            btnPolicyDocToV14PDF.Size = new System.Drawing.Size(114, 43);
            btnPolicyDocToV14PDF.TabIndex = 4;
            btnPolicyDocToV14PDF.Text = "Convert Policy Doc to V14 PDF";
            btnPolicyDocToV14PDF.UseCompatibleTextRendering = true;
            btnPolicyDocToV14PDF.UseVisualStyleBackColor = true;
            btnPolicyDocToV14PDF.Click += btnPolicyDocToV14PDF_Click;
            // 
            // btnRectifyPolicyDocs
            // 
            btnRectifyPolicyDocs.Location = new System.Drawing.Point(432, 84);
            btnRectifyPolicyDocs.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnRectifyPolicyDocs.Name = "btnRectifyPolicyDocs";
            btnRectifyPolicyDocs.Size = new System.Drawing.Size(104, 43);
            btnRectifyPolicyDocs.TabIndex = 5;
            btnRectifyPolicyDocs.Text = "Rectify Policy Docs";
            btnRectifyPolicyDocs.UseCompatibleTextRendering = true;
            btnRectifyPolicyDocs.UseVisualStyleBackColor = true;
            btnRectifyPolicyDocs.Click += btnRectifyPolicyDocs_Click;
            // 
            // cmbCustomer
            // 
            cmbCustomer.FormattingEnabled = true;
            cmbCustomer.Items.AddRange(new object[] { "CORESPECIALTY" });
            cmbCustomer.Location = new System.Drawing.Point(49, 84);
            cmbCustomer.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            cmbCustomer.Name = "cmbCustomer";
            cmbCustomer.Size = new System.Drawing.Size(138, 23);
            cmbCustomer.TabIndex = 6;
            // 
            // btnCopyToDestination
            // 
            btnCopyToDestination.Location = new System.Drawing.Point(542, 84);
            btnCopyToDestination.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnCopyToDestination.Name = "btnCopyToDestination";
            btnCopyToDestination.Size = new System.Drawing.Size(105, 43);
            btnCopyToDestination.TabIndex = 7;
            btnCopyToDestination.Text = "Copy To Destination";
            btnCopyToDestination.UseVisualStyleBackColor = true;
            btnCopyToDestination.Click += btnCopyToDestination_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(36, 53);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(647, 15);
            label1.TabIndex = 8;
            label1.Text = "________________________________________________________________________________________________________________________________";
            // 
            // txtPolicyTemplateList
            // 
            txtPolicyTemplateList.Location = new System.Drawing.Point(38, 187);
            txtPolicyTemplateList.Multiline = true;
            txtPolicyTemplateList.Name = "txtPolicyTemplateList";
            txtPolicyTemplateList.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            txtPolicyTemplateList.Size = new System.Drawing.Size(498, 316);
            txtPolicyTemplateList.TabIndex = 9;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(38, 129);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(647, 15);
            label2.TabIndex = 10;
            label2.Text = "________________________________________________________________________________________________________________________________";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(37, 163);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(129, 15);
            label3.TabIndex = 11;
            label3.Text = "Policy Template names";
            // 
            // btnCopyPolicyTemplates
            // 
            btnCopyPolicyTemplates.Location = new System.Drawing.Point(559, 187);
            btnCopyPolicyTemplates.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            btnCopyPolicyTemplates.Name = "btnCopyPolicyTemplates";
            btnCopyPolicyTemplates.Size = new System.Drawing.Size(105, 43);
            btnCopyPolicyTemplates.TabIndex = 12;
            btnCopyPolicyTemplates.Text = "Copy Policy Templates";
            btnCopyPolicyTemplates.UseVisualStyleBackColor = true;
            btnCopyPolicyTemplates.Click += btnCopyPolicyTemplates_Click;
            // 
            // TestForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(689, 543);
            Controls.Add(btnCopyPolicyTemplates);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(txtPolicyTemplateList);
            Controls.Add(label1);
            Controls.Add(btnCopyToDestination);
            Controls.Add(cmbCustomer);
            Controls.Add(btnRectifyPolicyDocs);
            Controls.Add(btnPolicyDocToV14PDF);
            Controls.Add(btnPolicyDocToV23PDF);
            Controls.Add(btnOnlyCompare);
            Controls.Add(btnDiffNCompare);
            Controls.Add(btnConvertToV23PDF);
            Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            Name = "TestForm";
            Text = "TestForm";
            Load += TestForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button btnConvertToV23PDF;
        private System.Windows.Forms.Button btnDiffNCompare;
        private System.Windows.Forms.Button btnOnlyCompare;
        private System.Windows.Forms.Button btnPolicyDocToV23PDF;
        private System.Windows.Forms.Button btnPolicyDocToV14PDF;
        private System.Windows.Forms.Button btnRectifyPolicyDocs;
        private System.Windows.Forms.ComboBox cmbCustomer;
        private System.Windows.Forms.Button btnCopyToDestination;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPolicyTemplateList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnCopyPolicyTemplates;
    }
}
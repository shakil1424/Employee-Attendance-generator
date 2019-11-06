﻿namespace ExcelConversion
{
  partial class Form1
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
      this.radioButtonFifthFLoor = new System.Windows.Forms.RadioButton();
      this.radioButtonSixthFloor = new System.Windows.Forms.RadioButton();
      this.buttonGenerate = new System.Windows.Forms.Button();
      this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
      this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
      this.buttonSave = new System.Windows.Forms.Button();
      this.textBoxFileUpload = new System.Windows.Forms.TextBox();
      this.textBoxSaveLocation = new System.Windows.Forms.TextBox();
      this.buttonConvertExcel = new System.Windows.Forms.Button();
      this.label1 = new System.Windows.Forms.Label();
      this.label2 = new System.Windows.Forms.Label();
      this.labelComplete = new System.Windows.Forms.Label();
      this.button1 = new System.Windows.Forms.Button();
      this.pictureBox1 = new System.Windows.Forms.PictureBox();
      ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
      this.SuspendLayout();
      // 
      // radioButtonFifthFLoor
      // 
      this.radioButtonFifthFLoor.AutoSize = true;
      this.radioButtonFifthFLoor.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.radioButtonFifthFLoor.Location = new System.Drawing.Point(181, 164);
      this.radioButtonFifthFLoor.Name = "radioButtonFifthFLoor";
      this.radioButtonFifthFLoor.Size = new System.Drawing.Size(102, 18);
      this.radioButtonFifthFLoor.TabIndex = 0;
      this.radioButtonFifthFLoor.TabStop = true;
      this.radioButtonFifthFLoor.Text = "Fifth Floor";
      this.radioButtonFifthFLoor.UseVisualStyleBackColor = true;
      // 
      // radioButtonSixthFloor
      // 
      this.radioButtonSixthFloor.AutoSize = true;
      this.radioButtonSixthFloor.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.radioButtonSixthFloor.Location = new System.Drawing.Point(317, 164);
      this.radioButtonSixthFloor.Name = "radioButtonSixthFloor";
      this.radioButtonSixthFloor.Size = new System.Drawing.Size(102, 18);
      this.radioButtonSixthFloor.TabIndex = 1;
      this.radioButtonSixthFloor.TabStop = true;
      this.radioButtonSixthFloor.Text = "Sixth Floor";
      this.radioButtonSixthFloor.UseVisualStyleBackColor = true;
      // 
      // buttonGenerate
      // 
      this.buttonGenerate.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.buttonGenerate.Location = new System.Drawing.Point(447, 230);
      this.buttonGenerate.Name = "buttonGenerate";
      this.buttonGenerate.Size = new System.Drawing.Size(103, 23);
      this.buttonGenerate.TabIndex = 2;
      this.buttonGenerate.Text = "Upload File";
      this.buttonGenerate.UseVisualStyleBackColor = true;
      this.buttonGenerate.Click += new System.EventHandler(this.buttonGenerate_Click);
      // 
      // openFileDialog1
      // 
      this.openFileDialog1.FileName = "openFileDialog1";
      // 
      // buttonSave
      // 
      this.buttonSave.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.buttonSave.Location = new System.Drawing.Point(447, 293);
      this.buttonSave.Name = "buttonSave";
      this.buttonSave.Size = new System.Drawing.Size(103, 23);
      this.buttonSave.TabIndex = 3;
      this.buttonSave.Text = "Save Location";
      this.buttonSave.UseVisualStyleBackColor = true;
      this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
      // 
      // textBoxFileUpload
      // 
      this.textBoxFileUpload.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.textBoxFileUpload.Location = new System.Drawing.Point(181, 233);
      this.textBoxFileUpload.Name = "textBoxFileUpload";
      this.textBoxFileUpload.Size = new System.Drawing.Size(210, 22);
      this.textBoxFileUpload.TabIndex = 4;
      // 
      // textBoxSaveLocation
      // 
      this.textBoxSaveLocation.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.textBoxSaveLocation.Location = new System.Drawing.Point(181, 295);
      this.textBoxSaveLocation.Name = "textBoxSaveLocation";
      this.textBoxSaveLocation.Size = new System.Drawing.Size(210, 22);
      this.textBoxSaveLocation.TabIndex = 5;
      // 
      // buttonConvertExcel
      // 
      this.buttonConvertExcel.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.buttonConvertExcel.Location = new System.Drawing.Point(309, 388);
      this.buttonConvertExcel.Name = "buttonConvertExcel";
      this.buttonConvertExcel.Size = new System.Drawing.Size(75, 23);
      this.buttonConvertExcel.TabIndex = 6;
      this.buttonConvertExcel.Text = "Convert";
      this.buttonConvertExcel.UseVisualStyleBackColor = true;
      this.buttonConvertExcel.Click += new System.EventHandler(this.buttonConvertExcel_Click);
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Font = new System.Drawing.Font("Consolas", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.label1.Location = new System.Drawing.Point(230, 44);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(271, 34);
      this.label1.TabIndex = 7;
      this.label1.Text = "Excel Converter ";
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.label2.Location = new System.Drawing.Point(448, 167);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(119, 15);
      this.label2.TabIndex = 8;
      this.label2.Text = "Choose The Floor";
      // 
      // labelComplete
      // 
      this.labelComplete.AutoSize = true;
      this.labelComplete.Font = new System.Drawing.Font("Consolas", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.labelComplete.ForeColor = System.Drawing.Color.DeepSkyBlue;
      this.labelComplete.Location = new System.Drawing.Point(286, 337);
      this.labelComplete.Name = "labelComplete";
      this.labelComplete.Size = new System.Drawing.Size(120, 18);
      this.labelComplete.TabIndex = 10;
      this.labelComplete.Text = "100% Completed";
      // 
      // button1
      // 
      this.button1.Image = global::ExcelConversion.Properties.Resources.resize_1572435079167937128closeicon7;
      this.button1.Location = new System.Drawing.Point(451, 381);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(41, 35);
      this.button1.TabIndex = 11;
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // pictureBox1
      // 
      this.pictureBox1.Image = global::ExcelConversion.Properties.Resources.ajax_loader;
      this.pictureBox1.InitialImage = global::ExcelConversion.Properties.Resources.Progress_Bar;
      this.pictureBox1.Location = new System.Drawing.Point(327, 337);
      this.pictureBox1.Name = "pictureBox1";
      this.pictureBox1.Size = new System.Drawing.Size(32, 32);
      this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
      this.pictureBox1.TabIndex = 9;
      this.pictureBox1.TabStop = false;
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.BackColor = System.Drawing.SystemColors.Window;
      this.ClientSize = new System.Drawing.Size(726, 450);
      this.Controls.Add(this.button1);
      this.Controls.Add(this.labelComplete);
      this.Controls.Add(this.pictureBox1);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.buttonConvertExcel);
      this.Controls.Add(this.textBoxSaveLocation);
      this.Controls.Add(this.textBoxFileUpload);
      this.Controls.Add(this.buttonSave);
      this.Controls.Add(this.buttonGenerate);
      this.Controls.Add(this.radioButtonSixthFloor);
      this.Controls.Add(this.radioButtonFifthFLoor);
      this.Name = "Form1";
      this.Text = "Convert Excel";
      this.Load += new System.EventHandler(this.Form1_Load);
      ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.RadioButton radioButtonFifthFLoor;
    private System.Windows.Forms.RadioButton radioButtonSixthFloor;
    private System.Windows.Forms.Button buttonGenerate;
    private System.Windows.Forms.OpenFileDialog openFileDialog1;
    private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    private System.Windows.Forms.Button buttonSave;
    private System.Windows.Forms.TextBox textBoxFileUpload;
    private System.Windows.Forms.TextBox textBoxSaveLocation;
    private System.Windows.Forms.Button buttonConvertExcel;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.PictureBox pictureBox1;
    private System.Windows.Forms.Label labelComplete;
    private System.Windows.Forms.Button button1;
  }
}

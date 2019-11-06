using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Floors;

namespace ExcelConversion
{
  public partial class Form1 : Form
  {
    private Thread newThreadForConversion;
    public Form1()
    {
      InitializeComponent();
      pictureBox1.Hide();
      labelComplete.Hide();

    }
    private void Form1_Load(object sender, EventArgs e)
    {
      //pictureBox1.Hide();

      labelComplete.Hide();
    }

    private void buttonGenerate_Click(object sender, EventArgs e)
    {
      OpenFileDialog openFileDialog1 = new OpenFileDialog
      {
        InitialDirectory = @"D:\",
        Title = "Select the excel file",
        CheckFileExists = true,
        CheckPathExists = true,
        RestoreDirectory = true,
        ReadOnlyChecked = true,
        ShowReadOnly = true
      };

      if (openFileDialog1.ShowDialog() == DialogResult.OK)
      {
        textBoxFileUpload.Text = openFileDialog1.FileName;     
      }
    }

    private void buttonSave_Click(object sender, EventArgs e)
    {
      var folderBrowserDialog1 = new FolderBrowserDialog();
      DialogResult result = folderBrowserDialog1.ShowDialog();
      if (result == DialogResult.OK)
      {
        string folderName = folderBrowserDialog1.SelectedPath;
        textBoxSaveLocation.Text = folderName;
      }
    }

    private void buttonConvertExcel_Click(object sender, EventArgs e)
    {
      pictureBox1.Show();
      var executeConversionProcess = true;
      string fileUpload = "";
      string saveLocation = "";
      if (textBoxFileUpload.Text == "")
      {
        MessageBox.Show("Select an excel file to upload");
        executeConversionProcess = false;
      }
      if (textBoxSaveLocation.Text == "")
      {
        MessageBox.Show("Select a save location");
        executeConversionProcess = false;
      }

      if (executeConversionProcess)
      {
        var sourceFileDirectory = textBoxFileUpload.Text;
        var generatedFilesDirectory = textBoxSaveLocation.Text;
       
        newThreadForConversion = new Thread(()=>
        {
          try
          {
            ConvertExcel(sourceFileDirectory, generatedFilesDirectory);
            labelComplete.Invoke((MethodInvoker)(() => labelComplete.Text = @"Operation Successful"));
          }
          catch
          {
            pictureBox1.Invoke((MethodInvoker)(() => pictureBox1.Hide()));
            labelComplete.Invoke((MethodInvoker)(() => labelComplete.Text = @"Operation unsuccessful"));
           
          }
          finally
          {
            pictureBox1.Invoke((MethodInvoker)(() => pictureBox1.Hide()));
            labelComplete.Invoke((MethodInvoker)(() => labelComplete.Show()));
          }
        });
        newThreadForConversion.Start();
      }      
    }
    private void ConvertExcel(string fileUpload, string saveLocation)
    {
      int choice=1;
      if (radioButtonFifthFLoor.Checked)
      {
        choice = 1;
      }
      else if(radioButtonSixthFloor.Checked)
      {
        choice = 2;
      }
      var floor = GetFloorInstance(choice);
      floor.CheckExcellProcesses();
      try
      {
        floor._ReadExcel = new Excel(@fileUpload, 1);
        floor._SaveLocation = @saveLocation;
        floor.CreateMultipleFiles();
        floor._ReadExcel.Close();
        floor._writeExcel.Close();
        //floor.KillExcel();
      }
      catch (Exception e)
      {
        MessageBox.Show(e.StackTrace);
        floor.KillExcel();
        

        
      }
      floor.KillExcel();
      
    }
    public FloorNo  GetFloorInstance(int choice)
    {
      if (choice == 1)
        return new FifthFloor();
      
      return new SixthFloor();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      Process[] AllProcesses = Process.GetProcessesByName("excel");
      
      foreach (Process ExcelProcess in AllProcesses)
      {

        ExcelProcess.Kill();
      }
      AllProcesses = null;
      if (newThreadForConversion!=null && newThreadForConversion.IsAlive)
      {
        try
        {
          newThreadForConversion.Abort();
        }
        catch (Exception ex)
        {
          //MessageBox.Show(ex.ExceptionState.ToString());
          
        }
       
      }
          
    }
  }
}

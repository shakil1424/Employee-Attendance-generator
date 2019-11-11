using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
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
      labelComplete.Hide();
      button1.Hide();
    }

    private void buttonGenerate_Click(object sender, EventArgs e)
    {
      labelComplete.Hide();
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
      labelComplete.Hide();
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
      var executeConversionProcess = true;
      string fileUpload = "";
      string saveLocation = "";
      if (textBoxFileUpload.Text == "")
      {
        MessageBox.Show("Select an excel file to upload", "Missing Source File", MessageBoxButtons.OK,
          MessageBoxIcon.Error);
        executeConversionProcess = false;
      }
      if (textBoxSaveLocation.Text == "")
      {
        MessageBox.Show("Select a save location", "Missing Location", MessageBoxButtons.OK, MessageBoxIcon.Error);
        executeConversionProcess = false;
      }
      if (textBoxFileUpload.Text.Contains("xls") && radioButtonSixthFloor.Checked)
      {
        MessageBox.Show("Choose appropriate excel file and floor ", "Incorrect input ", MessageBoxButtons.OK,
          MessageBoxIcon.Error);
        executeConversionProcess = false;
      }
      if (textBoxFileUpload.Text.Contains("dat") && radioButtonFifthFLoor.Checked)
      {
        MessageBox.Show("Choose appropriate excel file and floor ", "Incorrect input ", MessageBoxButtons.OK,
          MessageBoxIcon.Error);
        executeConversionProcess = false;
      }

      if (executeConversionProcess)
      {
        labelComplete.Hide();
        buttonConvertExcel.Hide();
        button1.Show();
        pictureBox1.Show();
        var sourceFileDirectory = textBoxFileUpload.Text;
        var generatedFilesDirectory = textBoxSaveLocation.Text;
        newThreadForConversion = new Thread(() =>
        {
          try
          {
            ConvertExcel(sourceFileDirectory, generatedFilesDirectory);
            labelComplete.Invoke((MethodInvoker) (() => labelComplete.Text = @"Operation Successful"));
          }
          catch (Exception ex)
          {
            pictureBox1.Invoke((MethodInvoker) (() => pictureBox1.Hide()));
            labelComplete.Invoke((MethodInvoker) (() => labelComplete.Text = @"Operation Unsuccessful"));
          }
          finally
          {
            pictureBox1.Invoke((MethodInvoker) (() => pictureBox1.Hide()));
            labelComplete.Invoke((MethodInvoker) (() => labelComplete.Show()));
            button1.Invoke((MethodInvoker) (() => button1.Hide()));
            buttonConvertExcel.Invoke((MethodInvoker) (() => buttonConvertExcel.Show()));
          }
        });
        newThreadForConversion.Start();
      }
    }
    private void ConvertExcel(string fileUpload, string saveLocation)
    {
      int choice = 1;
      if (radioButtonFifthFLoor.Checked)
      {
        choice = 1;
      }
      else if (radioButtonSixthFloor.Checked)
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
      }
      catch (Exception e)
      {
        WriteExceptionLog(e.StackTrace.ToString());
        floor.KillExcel();
      }
      floor.KillExcel();
    }
    public FloorNo GetFloorInstance(int choice)
    {
      if (choice == 1)
        return new FifthFloor();

      return new SixthFloor();
    }

    private async void button1_Click(object sender, EventArgs e)
    {
      pictureBox1.Hide();
      Process[] AllProcesses = Process.GetProcessesByName("excel");

      foreach (Process ExcelProcess in AllProcesses)
      {
        ExcelProcess.Kill();
      }

      AllProcesses = null;
      if (newThreadForConversion != null && newThreadForConversion.IsAlive)
      {
        try
        {
          await ThreadAbortAsync();
        }
        catch (Exception ex)
        {
          WriteExceptionLog(ex.StackTrace.ToString());
        }
      }

      buttonConvertExcel.Show();
      button1.Hide();
    }

    private async Task ThreadAbortAsync()
    {
      await Task.Run(() => newThreadForConversion.Abort());
    }

    private void WriteExceptionLog(string message)
    {
      var executionDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
      var fileName = DateTime.Today.ToString("dd-MM-yyyy") + ".txt";
      var filePath = executionDirectory + "\\" + fileName;
      using (var textWriter = new StreamWriter(filePath, true))
      {
        textWriter.WriteLine(message);
      }
    }
  }
}
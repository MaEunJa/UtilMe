using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace FindWordsInExcel
{
    public partial class Form1 : Form
    {
        private BackgroundWorker backgroundWorker;
        public Form1()
        {
            InitializeComponent();

            textSrc.Height = 38;
            textSrc.AutoSize = false; // 자동 크기 조절 비활성화
            textKeyword.Height = 38;
            textKeyword.AutoSize = false;

            // BackgroundWorker 초기화
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.DoWork += backgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged;

        }

        private void btnSrcFind_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "폴더를 선택하세요.";

            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string selectedFolder = folderBrowserDialog.SelectedPath;
                //Console.WriteLine("선택한 폴더 경로: " + selectedFolder);
                textSrc.Text = selectedFolder;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textColum.Text))
            {
                MessageBox.Show("탬색할 컬럼을 입력해주세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(textKeyword.Text))
            {
                MessageBox.Show("탬색할 단어를 입력해주세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(textSrc.Text))
            {
                MessageBox.Show("탬색할 폴더를 선택해주세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            backgroundWorker.RunWorkerAsync();



        }
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 검색 작업을 백그라운드 스레드에서 수행
            // progressBar1 값을 업데이트하여 진행 상황을 표시
            int colNum = int.Parse(textColum.Text);
            string result = "";
            DirectoryInfo directory = new DirectoryInfo(textSrc.Text);
            int completedCount = 0;
            progressBar1.Value = 0;

            int fileCount = 0;

            foreach (FileInfo file in directory.GetFiles("*.xlsx"))
            {
                // 무시할 잠금 파일 처리
                if (file.Name.StartsWith("~$")) continue;
                fileCount++;
            }

                if (InvokeRequired)
            {
                // 현재 스레드가 UI 스레드가 아니라면, UI 스레드에서 실행하도록 호출
                Invoke(new Action(() => lblProcess.Text = "0/" + fileCount));
            }
            else
            {
                // 현재 스레드가 이미 UI 스레드인 경우 바로 업데이트
                lblProcess.Text = "0/" + fileCount;
            }
            if (directory.Exists)
            {
                foreach (FileInfo file in directory.GetFiles("*.xlsx"))
                {
                    // 무시할 잠금 파일 처리
                    if (file.Name.StartsWith("~$")) continue;
                    SearchWordInExcelFile(file.FullName, textKeyword.Text, colNum);
                    completedCount++;
                    // 진행 상황 업데이트
                    int progress = (int)((double)completedCount / fileCount * 100);
                    backgroundWorker.ReportProgress(progress);

                    if (InvokeRequired)
                    {
                        // 현재 스레드가 UI 스레드가 아니라면, UI 스레드에서 실행하도록 호출
                        Invoke(new Action(() => lblProcess.Text = completedCount+"/" + fileCount));
                    }
                    else
                    {
                        // 현재 스레드가 이미 UI 스레드인 경우 바로 업데이트
                        lblProcess.Text = completedCount + "/" + fileCount;
                    }
                }
            }
            else
            {
                Console.WriteLine("폴더를 찾을 수 없습니다.");
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 백그라운드 스레드에서 보낸 진행 상황 값을 사용하여 ProgressBar 업데이트

            if (InvokeRequired)
            {
                // 현재 스레드가 UI 스레드가 아니라면, UI 스레드에서 실행하도록 호출
                Invoke(new Action(() => progressBar1.Value = e.ProgressPercentage));
            }
            else
            {
                // 현재 스레드가 이미 UI 스레드인 경우 바로 업데이트
                progressBar1.Value = e.ProgressPercentage;
            }
        }


        public string SearchWordInExcelFile(string filePath, string searchWord,int colNum)
        {
            string result = "";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            foreach (Excel.Range cell in range.Cells)
            {
                if (cell.Column == colNum && cell.Value != null && cell.Value.ToString().IndexOf(searchWord, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"파일: {Path.GetFileName(filePath)}, 찾은 위치: 행 {cell.Row}, 열 {cell.Column}");

                    if (InvokeRequired)
                    {
                        // 현재 스레드가 UI 스레드가 아니라면, UI 스레드에서 실행하도록 호출
                        Invoke(new Action(() =>
                        {
                            textResult.Text += $"파일: {Path.GetFileName(filePath)}, 찾은 위치: 행 {cell.Row}, 열 {cell.Column}\r\n";
                            textResult.SelectionStart = textResult.Text.Length; // 커서를 텍스트 끝으로 이동
                            textResult.ScrollToCaret(); // 스크롤을 커서 위치로 이동
                        }
                        ));
                    }
                    else
                    {
                        // 현재 스레드가 이미 UI 스레드인 경우 바로 업데이트
                        textResult.Text += $"파일: {Path.GetFileName(filePath)}, 찾은 위치: 행 {cell.Row}, 열 {cell.Column}\r\n";
                        textResult.SelectionStart = textResult.Text.Length; // 커서를 텍스트 끝으로 이동
                        textResult.ScrollToCaret(); // 스크롤을 커서 위치로 이동
                    }



                }
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            return result;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // 입력된 문자가 숫자나 제어 문자가 아닌 경우 입력을 막음
            }
        }
    }
}

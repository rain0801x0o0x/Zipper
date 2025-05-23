using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression; // ZIP機能に必要

namespace MySimpleZipper // プロジェクト作成時に指定した名前に合わせてください
{
    public partial class Form1 : Form
    {
        // UI コントロールの宣言
        private Label labelDropArea;
        private CheckedListBox checkedListBoxFiles;
        private Label labelZipNamePrompt;
        private TextBox textBoxZipFileName;
        private Button buttonCompress;

        // ドロップされたアイテムのパスを保持するリスト
        private List<string> droppedFilePaths = new List<string>();

        public Form1()
        {
            InitializeComponent();
            InitializeCustomUI();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        private void InitializeCustomUI()
        {
            this.Text = "簡易ZIP圧縮ツール (デスクトップOutputフォルダ保存)";
            this.Size = new Size(520, 480);
            this.MinimumSize = new Size(400, 350);

            labelZipNamePrompt = new Label();
            labelZipNamePrompt.Text = "保存するZIPファイル名 (拡張子不要):";
            labelZipNamePrompt.Dock = DockStyle.Bottom;
            labelZipNamePrompt.AutoSize = true;
            labelZipNamePrompt.Padding = new Padding(5, 0, 5, 2);

            textBoxZipFileName = new TextBox();
            textBoxZipFileName.Dock = DockStyle.Bottom;
            textBoxZipFileName.PlaceholderText = "例: backup_files";
            textBoxZipFileName.Padding = new Padding(5);

            buttonCompress = new Button();
            buttonCompress.Text = "選択アイテムを圧縮してOutputフォルダに保存";
            buttonCompress.Dock = DockStyle.Bottom;
            buttonCompress.Height = 40;
            buttonCompress.Padding = new Padding(5);
            buttonCompress.Click += new EventHandler(buttonCompress_Click);

            labelDropArea = new Label();
            labelDropArea.Text = "ここにファイルまたはフォルダを\r\nドラッグ＆ドロップしてください";
            labelDropArea.TextAlign = ContentAlignment.MiddleCenter;
            labelDropArea.Dock = DockStyle.Top;
            labelDropArea.Height = 100;
            labelDropArea.BorderStyle = BorderStyle.FixedSingle;
            labelDropArea.AllowDrop = true;
            labelDropArea.DragEnter += Form1_DragEnter;
            labelDropArea.DragDrop += Form1_DragDrop;
            labelDropArea.Visible = true;

            checkedListBoxFiles = new CheckedListBox();
            checkedListBoxFiles.Dock = DockStyle.Fill;
            checkedListBoxFiles.Visible = false;
            checkedListBoxFiles.CheckOnClick = true;
            checkedListBoxFiles.Padding = new Padding(5);

            this.Controls.Add(checkedListBoxFiles);
            this.Controls.Add(labelDropArea);
            this.Controls.Add(buttonCompress);
            this.Controls.Add(textBoxZipFileName);
            this.Controls.Add(labelZipNamePrompt);
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] droppedEntries = (string[])e.Data.GetData(DataFormats.FileDrop);
            bool itemAdded = false;
            foreach (string entryPath in droppedEntries)
            {
                if (!droppedFilePaths.Contains(entryPath))
                {
                    droppedFilePaths.Add(entryPath);
                    checkedListBoxFiles.Items.Add(entryPath, true);
                    itemAdded = true;
                }
            }

            if (itemAdded || droppedFilePaths.Count > 0)
            {
                labelDropArea.Visible = false;
                checkedListBoxFiles.Visible = true;
            }
        }

        private async void buttonCompress_Click(object sender, EventArgs e)
        {
            if (checkedListBoxFiles.CheckedItems.Count == 0)
            {
                MessageBox.Show("圧縮するファイルまたはフォルダを1つ以上選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string zipBaseName = textBoxZipFileName.Text.Trim();
            if (string.IsNullOrEmpty(zipBaseName))
            {
                MessageBox.Show("保存するZIPファイルのベース名を入力してください。\n例: my_archive", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxZipFileName.Focus();
                return;
            }
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                if (zipBaseName.Contains(invalidChar))
                {
                    MessageBox.Show($"ファイル名「{zipBaseName}」には使用できない文字「{invalidChar}」が含まれています。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBoxZipFileName.Focus();
                    return;
                }
            }

            // --- MODIFICATION START ---
            // 1. デスクトップパスを取得
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            // 2. Outputフォルダのパスを定義
            string outputFolderPath = Path.Combine(desktopPath, "Output");

            // 3. Outputフォルダが存在しない場合は作成
            try
            {
                if (!Directory.Exists(outputFolderPath))
                {
                    Directory.CreateDirectory(outputFolderPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"デスクトップにOutputフォルダを作成できませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 4. 完全なZIPファイルの保存パスを構築
            string zipPath = Path.Combine(outputFolderPath, zipBaseName + ".zip");
            // --- MODIFICATION END ---


            List<string> itemsToCompress = new List<string>();
            foreach (object item in checkedListBoxFiles.CheckedItems)
            {
                itemsToCompress.Add(item.ToString());
            }

            buttonCompress.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            try
            {
                await Task.Run(() => CreateZipFileProcess(zipPath, itemsToCompress));

                // メッセージを更新して保存場所を明確にする
                MessageBox.Show($"ZIPファイル '{Path.GetFileName(zipPath)}' が\nデスクトップのOutputフォルダに作成されました。", "圧縮完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                List<object> itemsToRemoveFromCheckedListBox = new List<object>();
                foreach (string compressedItemPath in itemsToCompress)
                {
                    droppedFilePaths.Remove(compressedItemPath);
                    for (int i = 0; i < checkedListBoxFiles.Items.Count; i++)
                    {
                        if (checkedListBoxFiles.Items[i].ToString() == compressedItemPath)
                        {
                            itemsToRemoveFromCheckedListBox.Add(checkedListBoxFiles.Items[i]);
                            break;
                        }
                    }
                }
                foreach (var itemToRemove in itemsToRemoveFromCheckedListBox)
                {
                    checkedListBoxFiles.Items.Remove(itemToRemove);
                }

                textBoxZipFileName.Clear();

                if (checkedListBoxFiles.Items.Count == 0)
                {
                    labelDropArea.Visible = true;
                    checkedListBoxFiles.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"圧縮中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                buttonCompress.Enabled = true;
                this.Cursor = Cursors.Default;
            }
        }

        private void CreateZipFileProcess(string zipPath, IEnumerable<string> itemsToCompress)
        {
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            using (ZipArchive archive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (string path in itemsToCompress)
                {
                    if (File.Exists(path))
                    {
                        archive.CreateEntryFromFile(path, Path.GetFileName(path));
                    }
                    else if (Directory.Exists(path))
                    {
                        AddDirectoryToZipArchive(archive, path, Path.GetFileName(path));
                    }
                }
            }
        }

        private void AddDirectoryToZipArchive(ZipArchive archive, string sourceDirectoryPath, string entryPrefixInZip)
        {
            foreach (string filePath in Directory.GetFiles(sourceDirectoryPath))
            {
                string entryName = Path.Combine(entryPrefixInZip, Path.GetFileName(filePath));
                archive.CreateEntryFromFile(filePath, entryName);
            }

            foreach (string subDirPath in Directory.GetDirectories(sourceDirectoryPath))
            {
                string nextEntryPrefix = Path.Combine(entryPrefixInZip, Path.GetFileName(subDirPath));
                AddDirectoryToZipArchive(archive, subDirPath, nextEntryPrefix);
            }
        }
    }
}
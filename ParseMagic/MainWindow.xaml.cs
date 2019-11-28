using Microsoft.Win32;
using Spire.Doc;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using Section = Spire.Doc.Section;

namespace ParseMagic
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        #region PROPS
        public int FileReciepientsCount
        {
            get { return Data.rowCount; }
            set
            {
                if (string.Equals(value, Data.rowCount))
                    return;
                Data.rowCount = value;
                OnPropertyChanged("FileReciepientsCount");
            }
        }

        public int FileAttributesCount
        {
            get { return Data.colCount; }
            set
            {
                if (string.Equals(value, Data.colCount))
                    return;
                Data.colCount = value;
                OnPropertyChanged("FileAttributesCount");
            }
        }

        public int ProcessedCount
        {
            get { return Data.processed; }
            set
            {
                if (string.Equals(value, Data.processed))
                    return;
                Data.processed = value;
                OnPropertyChanged("ProcessedCount");
            }
        }

        public int TextAttributesCount
        {
            get { return Data._textAttributeCount; }
            set
            {
                if (string.Equals(value, Data._textAttributeCount))
                    return;
                Data._textAttributeCount = value;
                OnPropertyChanged("TextAttributesCount");
            }
        }

        public string BaseTextStructure
        {
            get { return Data._baseTextStructure; }
            set
            {
                if (string.Equals(value, Data._baseTextStructure))
                    return;
                Data._baseTextStructure = value;
                OnPropertyChanged("BaseTextStructure");
            }
        }
        #endregion
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            TEXT_STRUCTURE_RTB.AppendText(Data._baseTextStructure);
            SetFoundTextAttributes(Data._baseTextStructure);
            populateSelectorComboBox();
        }

        private void populateSelectorComboBox()
        {
            foreach (string sel in Data.selectors)
            {
                ATT_SELECTOR_COMBO.Items.Add(sel);
            }
            ATT_SELECTOR_COMBO.SelectedIndex = 0;
        }

        private void SetFoundTextAttributes(string baseTextStructure)
        {
            var replace_parameters = baseTextStructure.Split('[', ']').Where((item, index) => index % 2 != 0).ToList();
            TextAttributesCount = replace_parameters.Count;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private async void LoadExcel(object sender, RoutedEventArgs e)
        {
            Data._type = "excel";
            Data.flush(); ProcessedCount = 0;
            OpenFileDialog dlg = new OpenFileDialog
            {

                // Set filter for file extension and default file extension 
                DefaultExt = ".xls",
                Filter = "All XCEL Files|*.xl;*.xls;*.xlsx;"
            };

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result.HasValue && result.Value)
            {
                string filename = dlg.FileName;
                Data._xcelfilelocation = filename;
                FILE_LOCATION_TB.Text = filename;

                SetApplicationBusy();
                await Task.Run(() => ParseExcelFile(filename));
                DATA_PEEK_DATAGRID.DataContext = Data.dt.DefaultView;
                SetApplicationIdle();

                if (FileAttributesCount == 0)
                {
                    MessageBox.Show("No attributes found in file");
                    setExportButtonState(false);
                }
                else
                {
                    setExportButtonState(true);
                }
            }
        }

        private void ParseExcelFile(string filename)
        {
            Data.ds = new DataSet();
            Data.dt = new DataTable();

            //Load the Excel file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Export data to data table
            Data.dt = sheet.ExportDataTable();
            Data.dt_as_array = Data.dt.AsEnumerable().Select(x => x.ItemArray).ToArray();

            FileReciepientsCount = Data.dt.Rows.Count;
            FileAttributesCount = Data.dt.Columns.Count;

            for (int c = 0; c < Data.colCount; c++)
            {
                Data.dt_as_array_columns.Add(Data.dt.Columns[c].ColumnName);
            }

            //Remove . in column header --- Gridview doesnt display data under column headers with a point
            for (int c = 0; c < Data.dt.Columns.Count; c++)
            {
                Data.dt.Columns[c].ColumnName = Data.dt.Columns[c].ColumnName.Replace(".", ",");
            }

            workbook.Dispose();
        }

        private async void LoadText(object sender, RoutedEventArgs e)
        {
            Data._type = "word";
            Data.flush(); ProcessedCount = 0;
            OpenFileDialog dlg = new OpenFileDialog
            {
                // Set filter for file extension and default file extension 
                DefaultExt = ".txt",
                Filter = "WORD AND TEXT Files|*.doc;*.docx;*.txt;"
            };

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result.HasValue && result.Value)
            {
                string filename = dlg.FileName;
                Data._textfilelocation = filename;
                FILE_LOCATION_TB.Text = filename;
                SetApplicationBusy();

                if (filename.Contains(".txt"))
                {
                    await Task.Run(() => ProcessText(ParseTextFile(filename)));
                }
                if (filename.Contains(".doc"))
                {
                    await Task.Run(() => ProcessText(ParseWordFile(filename)));
                }

                DATA_PEEK_DATAGRID.DataContext = new DataView(Data.dt);
                SetApplicationIdle();

                if (FileAttributesCount == 0)
                {
                    MessageBox.Show("No attributes found in file");
                    setExportButtonState(false);
                }
                else
                {
                    setExportButtonState(true);
                }
            }
        }

        private string[] ParseWordFile(string filename)
        {
            try
            {
                Document document = new Document();
                document.LoadFromFile(filename);

                StringBuilder sb = new StringBuilder();

                foreach (Section section in document.Sections)
                {
                    foreach (Spire.Doc.Documents.Paragraph paragraph in section.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Text);
                    }
                }

                document.Dispose();
                string outt = sb.ToString();
                return outt.Split('\r');
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("File in use by another Application.");
                return null;
            }
        }

        private string[] ParseTextFile(string filename)
        {
            return File.ReadAllLines(filename);
        }

        private void ProcessText(string[] data)
        {
            if (data != null)
            {
                Data.tuples = new List<List<string>>();
                List<string> tuple = new List<string>();

                for (int c = 0; c < data.Length; c++)
                {
                    string input = data[c];

                    try
                    {
                        Data.fileAttributesFound.Add(input.Split(Data.leftSelector, Data.rightSelector)[1]);
                    }
                    catch { }
                }

                int dataBodyStartIndex = 0;
                for (int c = 0; c < data.Length; c++)
                {
                    if (int.TryParse(data[c], out int number))
                    {
                        dataBodyStartIndex = c;
                        break;
                    }
                }

                for (int c = dataBodyStartIndex; c < data.Length; c++)
                {
                    if ((data[c].Length > 0) && data[c] != "\n")
                        tuple.Add(data[c].Replace("\n", ""));

                    if ((data[c] == "" || data[c] == "\n" || data[c] == " " || data[c] == "\r") && tuple.Count > 0)
                    {
                        Data.tuples.Add(new List<string>(tuple));
                        tuple.Clear();
                    }

                    //Thread.Sleep(500);
                    FileAttributesCount = Data.fileAttributesFound.Count;
                    FileReciepientsCount = Data.tuples.Count;
                }

                Data.ds = new DataSet();
                Data.dt = new DataTable();
                Data.dt.Clear();

                foreach (string s in Data.fileAttributesFound)
                {
                    Data.dt.Columns.Add(s.Replace(".", ","));
                }

                for (int c = 0; c < Data.tuples.Count; c++)
                {
                    DataRow row = Data.dt.NewRow();
                    foreach (string s in Data.fileAttributesFound)
                    {
                        row[s.Replace(".", ",")] = Data.tuples[c][Data.fileAttributesFound.IndexOf(s)];
                    }
                    Data.dt.Rows.Add(row);
                }
            }
        }

        private async void Export(object sender, RoutedEventArgs e)
        {
            if (Data._type != "")
            {
                SetApplicationBusy();

                await Task.Run(() =>
                {
                    string outputFile = DateTime.Now.ToString().Replace("/", "-").Replace(":", ".") + ".txt";
                    Data._baseTextStructure = new TextRange(TEXT_STRUCTURE_RTB.Document.ContentStart, TEXT_STRUCTURE_RTB.Document.ContentEnd).Text;
                    int? parameter_position_in_dataset_columns;
                    ProcessedCount = 0;
                    string processed_output = "";
                    var input = Data._baseTextStructure;
                    var replace_parameters = input.Split(Data.leftSelector, Data.rightSelector).Where((item, index) => index % 2 != 0).ToList();

                    if (Data._type == "excel")
                    {
                        for (int c = 0; c < Data.dt_as_array.Length; c++)
                        {
                            string row = Data._baseTextStructure;

                            foreach (string parameter in replace_parameters)
                            {
                                parameter_position_in_dataset_columns = Data.dt_as_array_columns.IndexOf(parameter);

                                if (parameter_position_in_dataset_columns < 0)
                                {
                                    MessageBox.Show("Malformed Attributes. Check text format and try again.");
                                    return;
                                }
                                else if (replace_parameters.Count == 0)
                                {
                                    MessageBox.Show("Malformed Attributes. Check text format and try again.");
                                    return;
                                }

                                string parameter_with_selector = Data.leftSelector + parameter + Data.rightSelector;
                                row = row.Replace(parameter_with_selector, Data.dt_as_array[c][(int)parameter_position_in_dataset_columns].ToString().Trim());
                            }

                            processed_output += row + "\n\n";
                            ProcessedCount++;
                            Thread.Sleep(Data.sleepTimes[new Random().Next(Data.sleepTimes.Count - 1)]);
                        }
                    }
                    else if (Data._type == "word")
                    {
                        for (int c = 0; c < Data.tuples.Count; c++)
                        {
                            string row = Data._baseTextStructure;

                            foreach (string parameter in replace_parameters)
                            {
                                parameter_position_in_dataset_columns = Data.fileAttributesFound.IndexOf(parameter);

                                if (parameter_position_in_dataset_columns < 0)
                                {
                                    MessageBox.Show("Malformed Attributes. Check text format and try again.");
                                    return;
                                }
                                else if (replace_parameters.Count == 0)
                                {
                                    MessageBox.Show("Malformed Attributes. Check text format and try again.");
                                    return;
                                }

                                string parameter_with_selector = Data.leftSelector + parameter + Data.rightSelector;
                                row = row.Replace(parameter_with_selector, Data.tuples[c][(int)parameter_position_in_dataset_columns].ToString().Trim());
                            }

                            processed_output += row + "\n\n";
                            ProcessedCount++;
                            Thread.Sleep(Data.sleepTimes[new Random().Next(Data.sleepTimes.Count - 1)]);
                        }
                    }

                    using (StreamWriter sw = new StreamWriter(outputFile, false))
                    {
                        sw.WriteLine(processed_output);
                    }

                    if (Data._open_after_export)
                    {
                        Process.Start(outputFile);
                    }
                });

                SetApplicationIdle();
            }
            else
            {
                MessageBox.Show("No file Selected");
            }
        }

        private void UpdateTextParameterCountCaller(object sender, System.Windows.Input.KeyEventArgs e)
        {
            UpdateTextParameterCount();
        }

        private void UpdateTextParameterCount()
        {
            string input = new TextRange(TEXT_STRUCTURE_RTB.Document.ContentStart, TEXT_STRUCTURE_RTB.Document.ContentEnd).Text;
            var replace_parameters = input.Split(Data.leftSelector, Data.rightSelector).Where((item, index) => index % 2 != 0).ToList();
            TextAttributesCount = replace_parameters.Count;
        }

        private void loopGif(object sender, RoutedEventArgs e)
        {
            BUSY_IMG.Position = new TimeSpan(0, 0, 1);
            BUSY_IMG.Play();
        }

        private void showBusy()
        {
            BUSY_IMG.Visibility = Visibility.Visible;
            APP_BUSY_TIP_TB.Visibility = Visibility.Visible;
        }

        private void hideBusy()
        {
            BUSY_IMG.Visibility = Visibility.Hidden;
            APP_BUSY_TIP_TB.Visibility = Visibility.Hidden;
        }

        private void SetApplicationIdle()
        {
            LOAD_EXCEL_BTN.IsEnabled = true;
            LOAD_TEXT_BTN.IsEnabled = true;
            EXPORT_TEXT_BTN.IsEnabled = true;
            hideBusy();
        }

        private void SetApplicationBusy()
        {
            LOAD_EXCEL_BTN.IsEnabled = false;
            LOAD_TEXT_BTN.IsEnabled = false;
            EXPORT_TEXT_BTN.IsEnabled = false;
            showBusy();
        }

        private void OpenAfterExportClicked(object sender, RoutedEventArgs e)
        {
            if ((bool)OPEN_AFTER_EXPORT_CHECKBOX.IsChecked == true)
                Data._open_after_export = true;
            else
                Data._open_after_export = false;
        }

        private void updateAttributeSelector(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Data.leftSelector = ATT_SELECTOR_COMBO.SelectedItem.ToString().Split(' ')[0].ToCharArray()[0];
            Data.rightSelector = ATT_SELECTOR_COMBO.SelectedItem.ToString().Split(' ')[1].ToCharArray()[0];

            UpdateTextParameterCount();
        }

        private void setExportButtonState(bool state)
        {
            EXPORT_TEXT_BTN.IsEnabled = state;
        }
    }
}

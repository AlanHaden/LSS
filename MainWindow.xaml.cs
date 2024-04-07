using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static LSS.Global;
using static LSS.PDF;
using System.Diagnostics;
using System.IO.IsolatedStorage;
using Microsoft.Win32;
using System.Reflection;
using System.Globalization;
using System.Windows.Data;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;
using PdfSharp.Drawing;
using System.Windows.Media.Media3D;
using System.Diagnostics.Eventing.Reader;
using System.Text;
using System.Threading;
using System.Security.Cryptography;

namespace LSS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {

        //ProductData DB connection
        SqlConnection cn;
        SqlCommand cmd;
        //SqlDataReader dr;
        SqlDataAdapter da;

        //Engineer DB connection
        SqlConnection Engcn;
        SqlCommand Engcmd;
        //SqlDataReader Engdr;
        SqlDataAdapter Engda;
        DataTable Engdt = new DataTable();

        private string selectedMan;
        private string selectedProd;
        private string modelHolder;

        private bool coDataSaved = false;
        private bool raycapOrWeidmullerSelected = false;

        //string[] selectedManufacturer = new string[10];
        //string[] selectedProduct = new string[10];
        //string[] selectedProductType = new string[10];

        public MainWindow()
        {
            InitializeComponent();
            this.PreviewMouseWheel += MainWindow_PreviewMouseWheel;
            Version currentVersion = Assembly.GetEntryAssembly().GetName().Version;
            versionLbl.Content = $"Version: {currentVersion}";
            //Check for restore flag
            if (CheckForRestoreFlag())
            {
                restoreFunction();
            }

            // Load saved company details for use in the Header
            LoadUserData();
            if (coDataSaved)
            {
                CoDataBar.Visibility = Visibility.Hidden;
            }
            else
            {
                CoDataBar.Visibility = Visibility.Visible;
            }
            coAddressL1P1.Text = tbAddL1.Text; coAddressL1P2.Text = tbAddL1.Text;
            coAddressL2P1.Text = tbAddL2.Text; coAddressL2P2.Text = tbAddL2.Text;
            coAddressL3P1.Text = tbAddL3.Text; coAddressL3P2.Text = tbAddL3.Text;
            coAddressL4P1.Text = tbAddL4.Text; coAddressL4P2.Text = tbAddL4.Text;
            coTelNumberP1.Text = tbTelNo.Text; coTelNumberP2.Text = tbTelNo.Text;


            selectedMan = "Raycap";
            Manu_combobox.SelectedItem = selectedMan;

            cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ProductData.mdf;Integrated Security=True");
            Engcn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\EngineersData.mdf;Integrated Security=True");
            try
            {
                cn.Open();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            SetEngineersCombobox();



            //Set manufacturers comboBox contents from table names
            List<string> manufacturers = new List<string>();
            cmd = new SqlCommand("SELECT * FROM SYS.TABLES", cn);
            da = new SqlDataAdapter(cmd);
            DataTable det = new DataTable();
            da.Fill(det);
            dgridview.ItemsSource = det.DefaultView;

            manufacturers.Add("");

            manufacturers.Add("No Device Present"); // Add "No Device Present" as the top item

            foreach (DataRow row in det.Rows)
            {
                manufacturers.Add((string)row[0]);
            }

            manufacturers.Sort((x, y) =>
            {
                if (x == "No Device Present")
                    return -1; // Place "No Device Present" at the beginning
                else if (y == "No Device Present")
                    return 1; // Place "No Device Present" at the beginning
                else
                    return x.CompareTo(y); // Sort other manufacturers alphabetically
            });
            //Manu_combobox.ItemsSource = manufacturers;
            Manu_combobox.ItemsSource = manufacturers.Skip(2).ToList();
            Manu_combobox.SelectedItem = "Raycap";

            cbMan01.ItemsSource = manufacturers; cbMan02.ItemsSource = manufacturers; cbMan03.ItemsSource = manufacturers; cbMan04.ItemsSource = manufacturers; cbMan05.ItemsSource = manufacturers;
            cbMan06.ItemsSource = manufacturers; cbMan07.ItemsSource = manufacturers; cbMan08.ItemsSource = manufacturers; cbMan09.ItemsSource = manufacturers; cbMan10.ItemsSource = manufacturers;

            GetAllModelRecord("Raycap");

            //----------- Engineers DB ---------------
            try
            {
                Engcn.Open();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            GetEngineerData();
            //----------------------------------------

            companyLogoP1.Source = imageUploadField.Source;
            companyLogoP2.Source = imageUploadField.Source;

        }

        private void dgridviewEng_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataGrid dataGrid = (DataGrid)sender;

            if (dataGrid.SelectedItem != null)
            {
                try
                {
                    string query = "SELECT * FROM TblEngineers WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(query, Engcn))
                    {
                        DataRowView dataRowView = (DataRowView)dataGrid.SelectedItem;
                        object firstField = dataRowView.Row[0];
                        cmd.Parameters.AddWithValue("@Id", firstField);

                        Engcn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Retrieve the values from the reader and assign them to the text boxes
                                txtEngId.Text = reader["Id"].ToString();
                                txtEngName.Text = reader["Name"].ToString();
                                txtEngELITester.Text = reader["MFTTester"].ToString();
                                txtEngELISN.Text = reader["MFTSN"].ToString();
                                txtEngSPDTester.Text = reader["SPDTester"].ToString();
                                txtEngSPDSN.Text = reader["SPDSN"].ToString();

                                // Retrieve the "Sig" value (image data) from the reader
                                byte[] imageData = reader["Sig"] as byte[];
                                if (imageData != null && imageData.Length > 0)
                                {
                                    // Convert the image data to a BitmapImage
                                    using (MemoryStream memoryStream = new MemoryStream(imageData))
                                    {
                                        BitmapImage bitmapImage = new BitmapImage();
                                        bitmapImage.BeginInit();
                                        bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                                        bitmapImage.StreamSource = memoryStream;
                                        bitmapImage.EndInit();

                                        // Set the image source of the imgEngSig field
                                        imgEngSig.Source = bitmapImage;
                                    }
                                }
                                else
                                {
                                    // If no image data is present, you can assign a default image or leave it empty
                                    imgEngSig.Source = null;
                                }
                                //btnEngUpdate.Visibility = Visibility.Visible;
                                //btnEngDelete.Visibility = Visibility.Visible;
                                //btnEngSave.Visibility = Visibility.Hidden;
                                btnEngEdit.IsEnabled = true;
                            }
                            else
                            {
                                MessageBox.Show("Engineer not found.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while retrieving engineer data: " + ex.Message);
                }
                finally
                {
                    if (Engcn.State == ConnectionState.Open)
                    {
                        Engcn.Close();
                    }
                }

            }
        }

        private void GetEngineerData()
        {
            try
            {
                Engcmd = new SqlCommand("SELECT Id, Name, MFTTester, MFTSN, SPDTester, SPDSN, Sig FROM TblEngineers", Engcn);
                Engda = new SqlDataAdapter(Engcmd);

                // Clear the DataTable before filling it
                Engdt.Clear();

                Engda.FillSchema(Engdt, SchemaType.Source); // Fill the DataTable schema without data

                // Add a new column to hold the display value of the signature
                if (!Engdt.Columns.Contains("SignatureDisplay"))
                {
                    Engdt.Columns.Add("SignatureDisplay", typeof(string));
                }

                Engda.Fill(Engdt); // Fill the DataTable with data

                // Loop through the rows of the DataTable
                foreach (DataRow row in Engdt.Rows)
                {
                    // Retrieve the byte array from the "Sig" column
                    byte[] imageData = row.Field<byte[]>("Sig");

                    // Set the "SignatureDisplay" column value to "None" if image data is null or empty
                    if (imageData == null || imageData.Length == 0)
                    {
                        row["SignatureDisplay"] = "None";
                    }
                    else
                    {
                        row["SignatureDisplay"] = "Signature on file";
                    }
                }

                dgridviewEng.ItemsSource = Engdt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while retrieving engineer data: " + ex.Message);
            }
            finally
            {
                if (Engcn.State == ConnectionState.Open)
                {
                    Engcn.Close();

                }
            }
        }

        private void btnEngSave_Click(object sender, RoutedEventArgs e)
        {
            if (txtEngId.Text != "")
            {
                SaveEngineerData();
                SetEngineersCombobox();
            }
            else
            {
                MessageBox.Show("An Engineer ID number is required");
                txtEngId.Focus();
            }
        }

        private void SaveEngineerData()
        {
            try
            {
                string QUERY = "INSERT INTO TblEngineers (Id, Name, MFTTester, MFTSN, SPDTester, SPDSN, Sig) VALUES (@Id, @Name, @MFTTester, @MFTSN, @SPDTester, @SPDSN, @Sig)";
                using (SqlCommand cmd = new SqlCommand(QUERY, Engcn))
                {
                    cmd.Parameters.AddWithValue("@Id", txtEngId.Text);
                    cmd.Parameters.AddWithValue("@Name", txtEngName.Text);
                    cmd.Parameters.AddWithValue("@MFTTester", txtEngELITester.Text);
                    cmd.Parameters.AddWithValue("@MFTSN", txtEngELISN.Text);
                    cmd.Parameters.AddWithValue("@SPDTester", txtEngSPDTester.Text);
                    cmd.Parameters.AddWithValue("@SPDSN", txtEngSPDSN.Text);

                    // Convert the image source to a byte array
                    byte[] imageData = null;
                    if (imgEngSig.Source is BitmapSource bitmapSource)
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            BitmapEncoder encoder = new PngBitmapEncoder();
                            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                            encoder.Save(memoryStream);
                            imageData = memoryStream.ToArray();
                        }
                    }

                    if (imageData != null)
                        cmd.Parameters.Add("@Sig", SqlDbType.VarBinary, -1).Value = imageData;
                    else
                        cmd.Parameters.Add("@Sig", SqlDbType.VarBinary, -1).Value = DBNull.Value;

                    Engcn.Open(); // Open the connection
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Engineer data saved successfully.");

                    Engcn.Close(); // Close the connection

                    if (imageData != null && imageData.Length > 0)
                    {
                        Engdt.Clear();
                        Engda.Fill(Engdt);
                        dgridviewEng.ItemsSource = Engdt.DefaultView;
                    }

                    // Clear the existing data in the DataTable
                    Engdt.Clear();
                    // Refresh the data source of the grid view
                    Engda.Fill(Engdt);
                    dgridviewEng.ItemsSource = Engdt.DefaultView;
                }

            }
            catch (SqlException ex)
            {
                if (ex.Number == 2627) // Check for specific error number for duplicate key violation
                {
                    MessageBox.Show("Duplicate Engeneer ID. Please enter a unique ID.");
                }
                else
                {
                    MessageBox.Show("An error occurred while saving engineer data: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while saving engineer data: " + ex.Message);
            }
            finally
            {
                if (Engcn.State == ConnectionState.Open)
                {
                    Engcn.Close();
                }
            }
        }
        private void SetEngineersCombobox()
        {
            //Set Engineers combobox contents from Names
            List<String> engineers = new List<String>();
            Engcmd = new SqlCommand("Select * FROM TblEngineers", Engcn);
            Engda = new SqlDataAdapter(Engcmd);
            DataTable Engdet = new DataTable();
            Engda.Fill(Engdet);

            engineers.Add("");

            foreach (DataRow row in Engdet.Rows)
            {
                engineers.Add((string)row[1]);
            }
            engineers.Sort();
            comboEngineer.ItemsSource = engineers;
        }

        private void btnEngCancel_Click(object sender, RoutedEventArgs e)
        {
            resetButtons();
        }

        private void resetButtons()
        {
            // Clear the values from the text boxes
            txtEngId.IsEnabled = false;
            txtEngName.IsEnabled = false;
            txtEngELITester.IsEnabled = false;
            txtEngELISN.IsEnabled = false;
            txtEngSPDTester.IsEnabled = false;
            txtEngSPDSN.IsEnabled = false;
            //imgEngSig.Source = null;
            btnEngUpdate.IsEnabled = false;
            btnEngDelete.IsEnabled = false;
            btnEngSave.IsEnabled = false;
            btnUploadSig.IsEnabled = false;
            btnEngCan.IsEnabled = false;
        }

        private void btnEngUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtEngId.Text))
            {
                MessageBox.Show("An ID number is required");
                txtEngId.Focus();
                return;
            }

            try
            {
                string query = "UPDATE TblEngineers SET Name = @Name, MFTTester = @MFTTester, MFTSN = @MFTSN, SPDTester = @SPDTester, SPDSN = @SPDSN, Sig = @Sig WHERE Id = @Id";

                using (SqlCommand cmd = new SqlCommand(query, Engcn))
                {
                    cmd.Parameters.AddWithValue("@Id", txtEngId.Text);
                    cmd.Parameters.AddWithValue("@Name", txtEngName.Text);
                    cmd.Parameters.AddWithValue("@MFTTester", txtEngELITester.Text);
                    cmd.Parameters.AddWithValue("@MFTSN", txtEngELISN.Text);
                    cmd.Parameters.AddWithValue("@SPDTester", txtEngSPDTester.Text);
                    cmd.Parameters.AddWithValue("@SPDSN", txtEngSPDSN.Text);

                    // Convert the image source to a byte array
                    byte[] imageData = null;
                    if (imgEngSig.Source is BitmapSource bitmapSource)
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            BitmapEncoder encoder = new PngBitmapEncoder();
                            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                            encoder.Save(memoryStream);
                            imageData = memoryStream.ToArray();
                        }
                    }

                    cmd.Parameters.AddWithValue("@Sig", imageData);

                    Engcn.Open(); // Open the connection
                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Engineer data updated successfully.");

                        // Refresh the data source of the grid view
                        GetEngineerData();
                    }
                    else
                    {
                        MessageBox.Show("No records were updated.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while updating engineer data: " + ex.Message);
            }
            finally
            {
                if (Engcn.State == ConnectionState.Open)
                {
                    Engcn.Close();
                }
            }
            resetButtons();
        }

        private void btnEngDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to delete this record?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (txtEngId.Text != "")
                    {
                        string id = txtEngId.Text;

                        string query = "DELETE FROM TblEngineers WHERE Id = @Id";

                        using (SqlCommand cmd = new SqlCommand(query, Engcn))
                        {
                            cmd.Parameters.AddWithValue("@Id", id);

                            Engcn.Open();
                            int rowsAffected = cmd.ExecuteNonQuery();
                            Engcn.Close();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Engineer data deleted successfully.");
                                GetEngineerData(); // Refresh the data grid
                                ClearTextBoxes();
                            }
                            else
                            {
                                MessageBox.Show("No engineer data found with the specified ID.");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("An ID number is required.");
                        txtEngId.Focus();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while deleting engineer data: " + ex.Message);
                }
                finally
                {
                    if (Engcn.State == ConnectionState.Open)
                    {
                        Engcn.Close();
                    }
                }
                ClearTextBoxes();
                SetEngineersCombobox();
                resetButtons();
            }
        }
        private void btnEngEdit_Click(object sender, RoutedEventArgs e)
        {
            txtEngId.IsEnabled = false;
            txtEngName.IsEnabled = true;
            txtEngELITester.IsEnabled = true;
            txtEngELISN.IsEnabled = true;
            txtEngSPDTester.IsEnabled = true;
            txtEngSPDSN.IsEnabled = true;

            btnEngUpdate.IsEnabled = true;
            btnEngDelete.IsEnabled = true;
            btnEngSave.IsEnabled = false;
            btnUploadSig.IsEnabled = true;
            btnEngCan.IsEnabled = true;
        }

        private void btnEngNew_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                Engcn.Open();

                // Get the highest Product ID
                SqlCommand getMaxIdCmd = new SqlCommand("SELECT MAX(Id) FROM TblEngineers", Engcn);
                int maxProductId = (int)getMaxIdCmd.ExecuteScalar();

                // Set the next Product ID to the relevant field
                int nextProductId = maxProductId + 1;
                ClearTextBoxes();
                txtEngName.IsEnabled = true;
                txtEngELITester.IsEnabled = true;
                txtEngELISN.IsEnabled = true;
                txtEngSPDTester.IsEnabled = true;
                txtEngSPDSN.IsEnabled = true;

                txtEngId.Text = nextProductId.ToString();
                btnEngCan.IsEnabled = true;
                btnEngSave.IsEnabled = true;
                btnUploadSig.IsEnabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while creating new record: " + ex.Message);
            }
            finally
            {
                if (Engcn.State == ConnectionState.Open)
                {
                    Engcn.Close();
                }
            }
        }

        private void ClearTextBoxes()
        {
            txtEngId.Text = "";
            txtEngName.Text = "";
            txtEngELITester.Text = "";
            txtEngELISN.Text = "";
            txtEngSPDTester.Text = "";
            txtEngSPDSN.Text = "";
            imgEngSig.Source = null;
        }


        private void btnUploadSig_Click(object sender, RoutedEventArgs e)
        {
            if (txtEngId.Text != "")
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png";

                if (openFileDialog.ShowDialog() == true)
                {
                    try
                    {
                        string imagePath = openFileDialog.FileName;
                        BitmapImage bitmap = new BitmapImage(new Uri(imagePath));

                        // Close the database connection if it's open
                        if (Engcn.State == ConnectionState.Open)
                            Engcn.Close();

                        // Update the imgEngSig.Source
                        imgEngSig.Source = bitmap;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occurred while uploading the signature image: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Select or create an engineer with an Id number first");
            }
        }

        private void Image_MouseUp(object sender, MouseButtonEventArgs e)
        {
            string webLink = "https://lightningsurgesolutions.co.uk/";

            try
            {
                // Open the web link in the default web browser
                Process.Start(webLink);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open web link: " + ex.Message);
            }
        }

        private List<string> GetAllModelRecord(string manReq)
        {
            if (manReq != "")
            {
                cmd = new SqlCommand("Select * from " + manReq, cn);
                da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                //dgridview1.ItemsSource = dt.DefaultView;

                List<String> models = new List<String>(dt.Rows.Count);

                foreach (DataRow row in dt.Rows)
                    models.Add((String)row["Product"]);

                return models;
            }
            else
            {
                List<String> models = new List<String>();
                return models;
            }
        }

        private void ComboBoxMan_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            ComboBox prodComboBox;
            string man = comboBox.SelectedItem as string;
            man = man.Trim();

            int index = Convert.ToInt32(comboBox.Name.Substring(5, 2));
            string spds = "tbSPDsum" + index.ToString("D2");
            TextBlock spdSumtextBox = FindName(spds) as TextBlock;
            formDataArray[index - 1, 2] = man;
            if (index > 9)
            {
                prodComboBox = FindName("cbProd" + index) as ComboBox;
            }
            else
            {
                prodComboBox = FindName("cbProd0" + index) as ComboBox;
            }
            if (man != "No Device Present" && man != "")
            {
                prodComboBox.ItemsSource = GetAllModelRecord(man);
            }
            else if (man == "No Device Present")
            {
                prodComboBox.ItemsSource = " ";

                spdSumtextBox.Text = formDataArray[index - 1, 2];
            }
            else if (man == "")
            {
                spdSumtextBox.Text = "";
            }

            if (man == "Raycap" || man == "Weidmuller")
            {
                raycapOrWeidmullerSelected = true;
            }
            else
            {
                raycapOrWeidmullerSelected = false;
            }
        }

        private void ComboBoxProd_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            int index = Convert.ToInt32(comboBox.Name.Substring(6, 2));
            string prod = (sender as ComboBox).SelectedItem as string;
            //MessageBox.Show("formDataArray[" + (index - 1) + ", 3]: " + formDataArray[index - 1, 3]);
            formDataArray[index - 1, 3] = prod;

            string spdt = "spdType" + index.ToString("D2");
            string spds = "tbSPDsum" + index.ToString("D2");
            string minm = "minmov" + index.ToString("D2");
            string maxm = "maxmov" + index.ToString("D2");
            string ming = "mingdt" + index.ToString("D2");
            string maxg = "maxgdt" + index.ToString("D2");
            TextBox spdTypetextBox = FindName(spdt) as TextBox;
            TextBlock spdSumtextBox = FindName(spds) as TextBlock;
            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;

            if (prod == null || prod.Length == 0)
            {
                formDataArray[index - 1, 4] = "";
                spdTypetextBox.Text = formDataArray[index - 1, 4];
                spdSumtextBox.Text = "";
                minmov.Text = "";
                maxmov.Text = "";
                mingdt.Text = "";
                maxgdt.Text = "";
                return;
            }

            //formDataArray[index - 1, 3] = prod;
            List<String> d = new List<String>();
            d = GetSelectedItemData(formDataArray[index - 1, 2], prod);
            formDataArray[index - 1, 4] = d[0];
            spdTypetextBox.Text = formDataArray[index - 1, 4];
            spdSumtextBox.Text = formDataArray[index - 1, 2] + " " + formDataArray[index - 1, 3];

            minmov.Text = d[1];
            maxmov.Text = d[2];
            mingdt.Text = d[3];
            maxgdt.Text = d[4];

            refreshValidation(index);

            if (raycapOrWeidmullerSelected)
            {
                //No idea what to do here!!  
            }
        }

        private void refreshValidation(int index)
        {
            string[] textBoxNames = { "tbl1pe", "tbl2pe", "tbl3pe", "tbnpe", "tbl1n", "tbl2n", "tbl3n" };

            List<TextBox> textBoxes = new List<TextBox>();

            // Find and add TextBoxes to the collection
            foreach (string textBoxName in textBoxNames)
            {
                string fullName = textBoxName + index.ToString("D2");
                TextBox textBox = FindName(fullName) as TextBox;

                if (textBox != null)
                {
                    textBoxes.Add(textBox);
                }
            }

            // Iterate through the TextBoxes
            foreach (TextBox textBox in textBoxes)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    // Store the current text value
                    string originalText = textBox.Text;

                    // Set the TextBox text to the same value
                    textBox.Text = originalText;

                    // Raise the TextChanged event manually
                    textBox.RaiseEvent(new TextChangedEventArgs(TextBox.TextChangedEvent, UndoAction.None)
                    {
                        // Pass the original TextBox as the sender
                        Source = textBox
                    });
                }
            }
        }

        //----------------------

        //--------- DB Explorerer -----------
        private void Manu_combobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedMan = (sender as ComboBox).SelectedItem as string;
            selectedMan = selectedMan.Trim();

            GetAllModelRecordDBExplorer(selectedMan);
            txtModel.Text = "Select model from table above";
            txtType.Text = "";
            txtMinMOV.Text = "";
            txtMaxMOV.Text = "";
            txtMinGDT.Text = "";
            txtMaxGDT.Text = "";
            txtStockCode.Text = "";
        }

        private List<string> GetAllModelRecordDBExplorer(string manReq)
        {
            if (manReq != "")
            {
                cmd = new SqlCommand("Select * from " + manReq, cn);
                da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dgridviewDBExplorer.ItemsSource = dt.DefaultView;

                List<String> models = new List<String>(dt.Rows.Count);

                foreach (DataRow row in dt.Rows)
                    models.Add((String)row["Product"]);

                return models;
            }
            else
            {
                List<String> models = new List<String>();
                return models;
            }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataGrid dataGrid = (DataGrid)sender;

            if (dataGrid.SelectedItem != null)
            {
                // Assuming the "Product" column is the first column (index 0)
                DataRowView rowView = (DataRowView)dataGrid.SelectedItem;
                string product = rowView["Product"].ToString();
                //MessageBox.Show(product);

                // Do something with the product string
                selectedProd = product;
                cmd = new SqlCommand("Select * from " + selectedMan + " where Product = '" + selectedProd + "'", cn);
                da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dgridview.ItemsSource = dt.DefaultView;

                GetSelectedItemData(selectedMan, selectedProd);
            }
        }

        private List<String> GetSelectedItemData(string man, string prod)
        {
            //MessageBox.Show("man = " + man + " prod = " + prod);
            List<String> itemData = new List<String>();
            string t = "N/A";
            if (!string.IsNullOrEmpty(prod))
            {
                cmd = new SqlCommand("Select * from " + man + " where Product = '" + prod + "'", cn);

                da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Object prodOb = dt.Rows[0][0];
                Object movMinOb = dt.Rows[0][1];
                Object movMaxOb = dt.Rows[0][2];
                Object gdtMinOb = dt.Rows[0][3];
                Object gdtMaxOb = dt.Rows[0][4];
                Object typeOb = dt.Rows[0][5];
                Object stockOb = dt.Rows[0][6];
                itemData.Add(typeOb.ToString());
                itemData.Add(movMinOb.ToString());
                itemData.Add(movMaxOb.ToString());
                itemData.Add(gdtMinOb.ToString());
                itemData.Add(gdtMaxOb.ToString());

                if (prodOb.ToString() != "")
                {
                    txtModel.Text = prodOb.ToString();
                    modelHolder = txtModel.Text;
                }
                else
                {
                    txtModel.Text = "No Product Selected";
                }

                if (movMinOb.ToString() != "")
                {
                    txtMinMOV.Text = movMinOb.ToString();
                }
                else
                {
                    txtMinMOV.Text = "N/A";
                }

                if (movMaxOb.ToString() != "")
                {
                    txtMaxMOV.Text = movMaxOb.ToString();
                }
                else
                {
                    txtMaxMOV.Text = "N/A";
                }

                if (gdtMinOb.ToString() != "")
                {
                    txtMinGDT.Text = gdtMinOb.ToString();
                }
                else
                {
                    txtMinGDT.Text = "N/A";
                }

                if (gdtMaxOb.ToString() != "")
                {
                    txtMaxGDT.Text = gdtMaxOb.ToString();
                }
                else
                {
                    txtMaxGDT.Text = "N/A";
                }

                if (typeOb.ToString() != "")
                {
                    t = typeOb.ToString();
                    txtType.Text = t;
                }
                else
                {
                    txtType.Text = "N/A";
                }

                if (stockOb.ToString() != "")
                {
                    txtStockCode.Text = stockOb.ToString();
                }
                else
                {
                    txtStockCode.Text = "N/A";
                }
            }

            return itemData;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string enteredProductName = txtModel.Text.Trim();
            List<String> currentModels = GetAllModelRecordDBExplorer(selectedMan);
            if (string.IsNullOrEmpty(enteredProductName))
            {
                MessageBox.Show("Model field is empty!");
                return;
            }

            foreach (string existingProductName in currentModels)
            {
                if (existingProductName.Equals(enteredProductName, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("This model name already exists!");
                    txtModel.Focus();
                    return;
                }
            }
            SaveInfo();

        }
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            string QUERY = "UPDATE " + selectedMan + " " +
            "SET MOVmin = @MOVmin, MOVmax = @MOVmax, GDTmin = @GDTmin, GDTmax = @GDTmax, Type = @Type, StockCode = @StockCode " +
            "WHERE Product = @Product";

            SqlCommand CMD = new SqlCommand(QUERY, cn);
            CMD.Parameters.AddWithValue("@Product", txtModel.Text);
            CMD.Parameters.AddWithValue("@MOVmin", txtMinMOV.Text);
            CMD.Parameters.AddWithValue("@MOVmax", txtMaxMOV.Text);
            CMD.Parameters.AddWithValue("@GDTmin", txtMinGDT.Text);
            CMD.Parameters.AddWithValue("@GDTmax", txtMaxGDT.Text);
            CMD.Parameters.AddWithValue("@Type", txtType.Text);
            CMD.Parameters.AddWithValue("@StockCode", txtStockCode.Text);
            CMD.ExecuteNonQuery();
            dgridviewDBExplorer.ItemsSource = null;
            GetAllModelRecordDBExplorer(selectedMan);
            revert();

            MessageBox.Show("Updated");
        }

        protected void SaveInfo()
        {
            string QUERY = "INSERT INTO " + selectedMan + " " +
            "(Product, MOVmin, MOVmax, GDTmin, GDTmax, Type, StockCode) " +
            "VALUES (@Product, @MOVmin, @MOVmax, @GDTmin, @GDTmax, @Type, @StockCode)";

            SqlCommand CMD = new SqlCommand(QUERY, cn);
            CMD.Parameters.AddWithValue("@Product", txtModel.Text);
            CMD.Parameters.AddWithValue("@MOVmin", txtMinMOV.Text);
            CMD.Parameters.AddWithValue("@MOVmax", txtMaxMOV.Text);
            CMD.Parameters.AddWithValue("@GDTmin", txtMinGDT.Text);
            CMD.Parameters.AddWithValue("@GDTmax", txtMaxGDT.Text);
            CMD.Parameters.AddWithValue("@Type", txtType.Text);
            CMD.Parameters.AddWithValue("@StockCode", txtStockCode.Text);
            CMD.ExecuteNonQuery();
            dgridviewDBExplorer.ItemsSource = null;
            GetAllModelRecordDBExplorer(selectedMan);
            revert();
            MessageBox.Show("Saved");
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to delete this record?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                string QUERY = "Delete from " + selectedMan + " " +
                "where Product = @Product";

                SqlCommand CMD = new SqlCommand(QUERY, cn);
                CMD.Parameters.AddWithValue("@Product", txtModel.Text);
                CMD.ExecuteNonQuery();
                dgridviewDBExplorer.ItemsSource = null;
                GetAllModelRecordDBExplorer(selectedMan);
                txtModel.Text = "Select model from table above";
                txtType.Text = "";
                txtMinMOV.Text = "";
                txtMaxMOV.Text = "";
                txtMinGDT.Text = "";
                txtMaxGDT.Text = "";
                txtStockCode.Text = "";
                revert();
                MessageBox.Show("Deleted");
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (txtModel.Text != "Select model from table above")
            {
                txtModel.IsEnabled = true;
                lblModel.IsEnabled = false;
                btnDelete.IsEnabled = true;
                btnSave.IsEnabled = false;
                btnUpdate.IsEnabled = true;
                btnCan.IsEnabled = true;
                txtMaxGDT.IsEnabled = true;
                txtMaxMOV.IsEnabled = true;
                txtMinGDT.IsEnabled = true;
                txtMinMOV.IsEnabled = true;
                txtType.IsEnabled = true;
                txtStockCode.IsEnabled = true;
            }
        }

        private void btnCan_Click(object sender, RoutedEventArgs e)
        {
            revert();
        }

        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            txtModel.Text = "";
            txtType.Text = "";
            txtMinMOV.Text = "";
            txtMaxMOV.Text = "";
            txtMinGDT.Text = "";
            txtMaxGDT.Text = "";
            txtStockCode.Text = "";

            txtModel.IsEnabled = true;
            lblModel.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSave.IsEnabled = true;
            btnUpdate.IsEnabled = false;
            btnCan.IsEnabled = true;
            txtMaxGDT.IsEnabled = true;
            txtMaxMOV.IsEnabled = true;
            txtMinGDT.IsEnabled = true;
            txtMinMOV.IsEnabled = true;
            txtType.IsEnabled = true;
            txtStockCode.IsEnabled = true;
            txtModel.Focus();
        }

        private void txtModel_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtModel.IsVisible && txtModel.IsEnabled)
            {
                if (txtModel.Text == modelHolder)
                {
                    btnUpdate.IsEnabled = true;
                    btnSave.IsEnabled = false;
                }
                else if (string.IsNullOrEmpty(txtModel.Text) || txtModel.Text != modelHolder)
                {
                    btnSave.IsEnabled = true;
                    btnUpdate.IsEnabled = false;
                }
            }
        }

        private void revert()
        {
            txtModel.IsEnabled = false;
            lblModel.Visibility = Visibility.Hidden;
            btnDelete.IsEnabled = false;
            btnSave.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnCan.IsEnabled = false;
            txtMaxGDT.IsEnabled = false;
            txtMaxMOV.IsEnabled = false;
            txtMinGDT.IsEnabled = false;
            txtMinMOV.IsEnabled = false;
            txtType.IsEnabled = false;
            txtStockCode.IsEnabled = false;
        }

        private void tbsourceProtected_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(17, 2));

            string sps = "tbsourceProtectedsum" + index.ToString("D2");
            TextBlock sourceProtectedTextBoxsum = FindName(sps) as TextBlock;

            sourceProtectedTextBoxsum.Text = textBox.Text;
            formDataArray[index - 1, 0] = textBox.Text;

        }

        private void tbCircuitNo_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(11, 2));
            formDataArray[index - 1, 1] = textBox.Text;
        }

        //-------------------------------------------

        private void Tbl1pe_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 6] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "lpe";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl1pemg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 5] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 5].Substring(0, 3) == "mov") || (formDataArray[index - 1, 5].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 5].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }

        private void Tbl2pe_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 8] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "lpe";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl2pemg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 7] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 7].Substring(0, 3) == "mov") || (formDataArray[index - 1, 7].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 7].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }

        private void Tbl3pe_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 10] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "lpe";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl3pemg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 9] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 9].Substring(0, 3) == "mov") || (formDataArray[index - 1, 9].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 9].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }
        private void Tbnpe_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 12] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "npe";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbnpemg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 11] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 11].Substring(0, 3) == "mov") || (formDataArray[index - 1, 11].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 11].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }
        private void Tbl1n_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 14] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "ln";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl1nmg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 13] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 13].Substring(0, 3) == "mov") || (formDataArray[index - 1, 13].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 13].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }
        private void Tbl2n_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 16] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "ln";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl2nmg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 15] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 15].Substring(0, 3) == "mov") || (formDataArray[index - 1, 15].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 15].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }
        private void Tbl3n_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 18] = textBox.Text;
            string localManufacturer = formDataArray[index - index, 2];
            string pos = "ln";

            int ri = index;
            string minm = "minmov" + ri.ToString("D2");
            string maxm = "maxmov" + ri.ToString("D2");
            string ming = "mingdt" + ri.ToString("D2");
            string maxg = "maxgdt" + ri.ToString("D2");
            string mg = "tbl3nmg" + ri.ToString("D2");

            TextBlock minmov = FindName(minm) as TextBlock;
            TextBlock maxmov = FindName(maxm) as TextBlock;
            TextBlock mingdt = FindName(ming) as TextBlock;
            TextBlock maxgdt = FindName(maxg) as TextBlock;
            TextBlock movgdt = FindName(mg) as TextBlock;

            formDataArray[index - 1, 17] = ValidateData(textBox, localManufacturer, pos, minmov?.Text, maxmov?.Text, mingdt?.Text, maxgdt?.Text);
            if ((formDataArray[index - 1, 17].Substring(0, 3) == "mov") || (formDataArray[index - 1, 17].Substring(0, 3) == "gdt"))
            {
                movgdt.Text = formDataArray[index - 1, 17].Substring(0, 3);
            }
            else
            {
                movgdt.Text = "";
            }
        }
        private string ValidateData(TextBox valToBeTested, string manufacturer, string position, string movMin, string movMax, string gdtMin, string gdtMax)
        {
            if ((manufacturer != "Dehn") && (!string.IsNullOrEmpty(manufacturer)))
            {
                if (string.IsNullOrEmpty(valToBeTested?.Text) || !int.TryParse(valToBeTested.Text, out int value))
                {
                    return "invalid";
                }

                bool isMovSupplied = !string.IsNullOrEmpty(movMin) && !string.IsNullOrEmpty(movMax);
                bool isGdtSupplied = !string.IsNullOrEmpty(gdtMin) && !string.IsNullOrEmpty(gdtMax);

                if (!isMovSupplied && !isGdtSupplied)
                {
                    valToBeTested.Foreground = Brushes.Black;
                    return "no_data";
                }

                if (isMovSupplied && !isGdtSupplied)
                {
                    if (!int.TryParse(movMin, out int movMinValue) || !int.TryParse(movMax, out int movMaxValue))
                    {
                        return "invalid";
                    }

                    if (value < movMinValue || value > movMaxValue)
                    {
                        valToBeTested.Foreground = Brushes.Red;
                        return "result_OR";
                    }
                    else
                    {
                        valToBeTested.Foreground = Brushes.Black;
                        return "mov";
                    }
                }

                if (isGdtSupplied && !isMovSupplied)
                {
                    if (!int.TryParse(gdtMin, out int gdtMinValue) || !int.TryParse(gdtMax, out int gdtMaxValue))
                    {
                        return "invalid";
                    }

                    if (value < gdtMinValue || value > gdtMaxValue)
                    {
                        valToBeTested.Foreground = Brushes.Red;
                        return "result_OR";
                    }
                    else
                    {
                        valToBeTested.Foreground = Brushes.Black;
                        return "gdt";
                    }
                }

                if (isMovSupplied && isGdtSupplied)
                {
                    if (!int.TryParse(movMin, out int movMinValue) || !int.TryParse(movMax, out int movMaxValue) ||
                        !int.TryParse(gdtMin, out int gdtMinValue) || !int.TryParse(gdtMax, out int gdtMaxValue))
                    {
                        return "invalid";
                    }

                    if ((position != "lpe"))
                    {
                        if (value >= gdtMinValue && value <= gdtMaxValue)
                        {
                            valToBeTested.Foreground = Brushes.Black;
                            return "gdt";
                        }
                        else
                        {
                            valToBeTested.Foreground = Brushes.Red;
                            return "result_OR";
                        }
                    }
                    else
                    {
                        if (value >= movMinValue && value <= movMaxValue)
                        {
                            valToBeTested.Foreground = Brushes.Black;
                            return "mov";
                        }
                        else if (value >= gdtMinValue && value <= gdtMaxValue)
                        {
                            valToBeTested.Foreground = Brushes.Black;
                            return "gdt";
                        }
                        else
                        {
                            valToBeTested.Foreground = Brushes.Red;
                            return "result_OR";
                        }
                    }
                }
            }

            return "invalid";
        }

        private void Tbzs_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            int index = Convert.ToInt32(textBox.Name.Substring(4, 2));
            formDataArray[index - 1, 19] = textBox.Text;
        }

        //---------- 2nd table

        private void tbLocation_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(10, 2));
            string l = "tbLocationsum" + index.ToString("D2");
            TextBlock locationsum = FindName(l) as TextBlock;

            locationsum.Text = textBox.Text;
            formDataArray[index - 1, 20] = textBox.Text;

        }

        private void txtBoxReference_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            formDetails[0] = textBox.Text;
        }

        private void txtBoxCertNum_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            certNumP1.Content = textBox.Text;
            certNumP2.Content = textBox.Text;
            formDetails[1] = textBox.Text;
        }

        private void txtSiteAddress_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            formDetails[2] = textBox.Text;
        }

        private void comboEngineer_DropDownClosed(object sender, EventArgs e)
        {
            string selectedEngineer = (sender as ComboBox).SelectedItem as string;
            selectedEngineer = selectedEngineer.Trim();
            formDetails[3] = selectedEngineer;
            handleEngineerChanged();
        }

        private void handleEngineerChanged()
        {
            string selectedEngineer = formDetails[3];
            if (selectedEngineer != null)
            {

                Engcmd = new SqlCommand("Select * from TblEngineers where Name = '" + selectedEngineer + "'", Engcn);
                Engda = new SqlDataAdapter(Engcmd);

                // Create a DataTable to hold the retrieved data
                DataTable engineerData = new DataTable();
                Engda.Fill(engineerData);

                // Check if any rows were retrieved
                if (engineerData.Rows.Count > 0)
                {
                    DataRow engineerRow = engineerData.Rows[0];

                    // Save to array
                    formDetails[4] = engineerRow["MFTTester"].ToString();
                    formDetails[5] = engineerRow["MFTSN"].ToString();
                    formDetails[6] = engineerRow["SPDTester"].ToString();
                    formDetails[7] = engineerRow["SPDSN"].ToString();

                    // Populate the textboxes
                    tbSelEngMFTTester.Text = formDetails[4];
                    tbSelEngMFTSN.Text = formDetails[5];
                    tbSelEngSPDTester.Text = formDetails[6];
                    tbSelEngSPDSN.Text = formDetails[7];

                    // Populate the image box with the signature
                    byte[] imageData = engineerRow.Field<byte[]>("Sig");
                    if (imageData != null && imageData.Length > 0)
                    {
                        using (MemoryStream memoryStream = new MemoryStream(imageData))
                        {
                            BitmapImage bitmapImage = new BitmapImage();
                            bitmapImage.BeginInit();
                            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                            bitmapImage.StreamSource = memoryStream;
                            bitmapImage.EndInit();
                            imgSelEngSignature.Source = bitmapImage;

                            BitmapEncoder encoder = new JpegBitmapEncoder();
                            encoder.Frames.Add(BitmapFrame.Create(bitmapImage));
                            encoder.Save(memoryStream);
                            imageSignatureData = memoryStream.ToArray();
                        }
                    }
                    else
                    {
                        // Clear the image box if no signature is present
                        imgSelEngSignature.Source = null;
                    }
                }
                else
                {
                    // Clear the array if no data found
                    // Save to array
                    formDetails[4] = "";
                    formDetails[5] = "";
                    formDetails[6] = "";
                    formDetails[7] = "";

                    // Clear the textboxes and image box if no data is found
                    tbSelEngMFTTester.Text = string.Empty;
                    tbSelEngMFTSN.Text = string.Empty;
                    tbSelEngSPDTester.Text = string.Empty;
                    tbSelEngSPDSN.Text = string.Empty;
                    imgSelEngSignature.Source = null;
                }
            }
        }

        private void txtDate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DatePicker dateBox = sender as DatePicker;
            formDetails[8] = dateBox.Text;
            string date_string = formDetails[8];
            DateTime date;

            if (DateTime.TryParseExact(date_string, "dd MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                DateTime new_date = date.AddYears(1);
                dueDate = new_date.ToString("dd MMMM yyyy");
                tbDueDate.Text = " " + dueDate;
            }
            else
            {
                dueDate = "Invalid date";
            }
        }

        private void btnCreatePDF_Click(object sender, RoutedEventArgs e)
        {
            if (formDetails[1] != null)
            {
                PrintToPDF();
            }
            else
            {
                MessageBox.Show("A certificate number MUST be entered");
            }
        }

        private void tbBSEN_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 21] = textBox.Text;
        }

        private void tbisotype_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(9, 2));
            formDataArray[index - 1, 22] = textBox.Text;
        }
        private void tbRating_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(8, 2));
            formDataArray[index - 1, 23] = textBox.Text;
        }
        private void tbSSC_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 24] = textBox.Text;
        }
        private void tbRefMeth_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(9, 2));
            formDataArray[index - 1, 25] = textBox.Text;
        }
        private void tbLive_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 26] = textBox.Text;
        }
        private void tbCPC_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(5, 2));
            formDataArray[index - 1, 27] = textBox.Text;
        }
        private void tbR2_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(4, 2));
            formDataArray[index - 1, 28] = textBox.Text;
        }
        private void tbIRLL_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(6, 2));
            formDataArray[index - 1, 29] = textBox.Text;
        }
        private void tbIRLPE_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int index = Convert.ToInt32(textBox.Name.Substring(7, 2));
            formDataArray[index - 1, 30] = textBox.Text;
        }
        private void cbP_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox textBox = sender as CheckBox;
            int index = Convert.ToInt32(textBox.Name.Substring(3, 2));
            formDataArray[index - 1, 31] = "True";
        }
        private void cbP_UnChecked(object sender, RoutedEventArgs e)
        {
            CheckBox textBox = sender as CheckBox;
            int index = Convert.ToInt32(textBox.Name.Substring(3, 2));
            formDataArray[index - 1, 31] = "";
        }
        private void cbP_Indeterminate(object sender, RoutedEventArgs e)
        {
            CheckBox textBox = sender as CheckBox;
            int index = Convert.ToInt32(textBox.Name.Substring(3, 2));
            formDataArray[index - 1, 31] = "False";
        }

        private void PassCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            string comments = "No remedial action required";

            CheckBox passCheckBox = (CheckBox)sender;
            StackPanel selectedGroup = FindParent<StackPanel>(passCheckBox);
            int index = Convert.ToInt32(passCheckBox.Name.Substring(12, 2));

            string statusbox = "tbStatus" + index.ToString("D2");
            TextBlock statusboxtb = FindName(statusbox) as TextBlock;
            statusboxtb.Foreground = Brushes.Black;
            statusboxtb.Text = "PASS";
            formDataArray[index - 1, 32] = comments;
            formDataArray[index - 1, 33] = "PASS";

            if (selectedGroup != null)
            {
                foreach (CheckBox checkBox in selectedGroup.Children.OfType<CheckBox>())
                {
                    if (checkBox != passCheckBox)
                    {
                        checkBox.IsChecked = false;
                    }
                }
            }
            checkOverAllStatus();
        }

        private void PassCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox passCheckBox = (CheckBox)sender;
            StackPanel selectedGroup = FindParent<StackPanel>(passCheckBox);
            int index = Convert.ToInt32(passCheckBox.Name.Substring(12, 2));

            string statusbox = "tbStatus" + index.ToString("D2");
            TextBlock statusboxtb = FindName(statusbox) as TextBlock;
            if (AreAllUnchecked(selectedGroup))
            {
                statusboxtb.Text = "";
                formDataArray[index - 1, 32] = "";
                formDataArray[index - 1, 33] = "";
            }
            checkOverAllStatus();
        }

        private bool AreAllUnchecked(StackPanel emptyBoxes)
        {
            foreach (var item in emptyBoxes.Children)
            {
                if (item is CheckBox checkBox && checkBox.IsChecked == true)
                {
                    return false; // Return false as soon as we find a checked checkbox
                }
            }

            // If we've gone through all checkboxes and haven't returned yet, then all checkboxes are unchecked
            return true;
        }

        private void NumberedCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox numberedCheckBox = (CheckBox)sender;
            StackPanel selectedGroup = FindParent<StackPanel>(numberedCheckBox);
            int index = Convert.ToInt32(selectedGroup.Name.Substring(16, 2));
            string statusbox = "tbStatus" + index.ToString("D2");
            TextBlock statusboxtb = FindName(statusbox) as TextBlock;

            if (AreAllUnchecked(selectedGroup))
            {
                statusboxtb.Text = "";
                formDataArray[index - 1, 32] = "";
                formDataArray[index - 1, 33] = "";
            }

            checkOverAllStatus();
        }

        private void NumberedCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            string comments = "";
            CheckBox numberedCheckBox = (CheckBox)sender;
            StackPanel selectedGroup = FindParent<StackPanel>(numberedCheckBox);
            int index = Convert.ToInt32(selectedGroup.Name.Substring(16, 2));
            string boxname = "PassCheckBox" + index.ToString("D2");
            CheckBox passCheckBox = selectedGroup?.FindName(boxname) as CheckBox;

            bool isNumberedChecked = false;
            List<string> checkedCodes = new List<string>();

            foreach (CheckBox checkBox in selectedGroup.Children.OfType<CheckBox>())
            {
                if (checkBox.IsChecked == true)
                {
                    string checkBoxContent = checkBox.Content.ToString();

                    if (checkBoxContent == "1")
                    {
                        isNumberedChecked = true;
                    }
                    else if (checkBoxContent == "2" || checkBoxContent == "3" || checkBoxContent == "4" || checkBoxContent == "5")
                    {
                        checkedCodes.Add(checkBoxContent);
                    }
                }
            }

            if (isNumberedChecked && checkedCodes.Count > 0)
            {
                comments = "C2 - code(s) " + string.Join(", ", checkedCodes) + " & C3 - code 1";
            }
            else if (isNumberedChecked)
            {
                comments = "C3 - code 1";
            }
            else if (checkedCodes.Count > 0)
            {
                comments = "C2 - code(s) " + string.Join(", ", checkedCodes);
            }



            string statusbox = "tbStatus" + index.ToString("D2");
            TextBlock statusboxtb = FindName(statusbox) as TextBlock;
            statusboxtb.Foreground = Brushes.Red;
            statusboxtb.Text = "FAIL";
            formDataArray[index - 1, 32] = comments;
            formDataArray[index - 1, 33] = "FAIL";

            if (passCheckBox != null)
            {
                if (passCheckBox.IsChecked == true)
                {
                    passCheckBox.IsChecked = false;
                }
            }

            checkOverAllStatus();

        }

        public void checkOverAllStatus()
        {
            bool allEmpty = true; // To check if all TextBlocks are empty
            formDetails[9] = "PASS";
            tbOverall.Foreground = Brushes.Black;
            tbOverall.Text = "PASS";

            for (int i = 0; i < 10; i++)
            {
                string statusbox = "tbStatus" + i.ToString("D2");
                TextBlock statusboxtb = FindName(statusbox) as TextBlock;

                if (statusboxtb != null)
                {
                    if (statusboxtb.Text == "FAIL")
                    {
                        formDetails[9] = "FAIL";
                        tbOverall.Foreground = Brushes.Red;
                        tbOverall.Text = "FAIL";
                    }
                    if (!string.IsNullOrWhiteSpace(statusboxtb.Text))
                    {
                        allEmpty = false; // Found a non-empty TextBlock
                    }
                }
            }

            if (allEmpty)
            {
                formDetails[9] = "";
                tbOverall.Text = "";
            }
        }

        // Helper method to find the parent of a given type in the visual tree
        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);

            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }

            return parent as T;
        }

        private void Tabby_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tabControl = (TabControl)sender;
            TabItem selectedTab = (TabItem)tabControl.SelectedItem;

            if (selectedTab != null && selectedTab.Header.ToString() == "Test Certificate Page 2")
            {
                liveGrid.Visibility = Visibility.Visible;
            }
            else
            {
                liveGrid.Visibility = Visibility.Hidden;
            }
        }

        private void btnSaveCompanyInfo_Click(object sender, RoutedEventArgs e)
        {
            SaveUserData();
        }

        // Save user data to local storage
        private void SaveUserData()
        {
            // Collect the data from the textboxes
            string[] textBoxValues = { tbAddL1.Text, tbAddL2.Text, tbAddL3.Text, tbAddL4.Text, tbTelNo.Text };
            string data = string.Join("|", textBoxValues.Select(value => value.Replace(",", "&#44;")));

            try
            {
                // Declare isolatedStorage outside the using block
                IsolatedStorageFile isolatedStorage = null;

                // Open or create the isolated storage file
                try
                {
                    isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly();

                    // Save the user data to a temporary file
                    string tempFilePath = "tempUserData.txt";
                    using (var stream = new StreamWriter(new IsolatedStorageFileStream(tempFilePath, FileMode.Create, isolatedStorage)))
                    {
                        // Write the user data to the temporary file
                        stream.WriteLine(data);
                    }

                    // Delete the existing user data file (if it exists)
                    string userDataFilePath = "UserData.txt";
                    if (isolatedStorage.FileExists(userDataFilePath))
                    {
                        isolatedStorage.DeleteFile(userDataFilePath);
                    }

                    // Rename the temporary file to the user data file
                    isolatedStorage.MoveFile(tempFilePath, userDataFilePath);

                    // Save the uploaded image to isolated storage if available
                    if (imageUploadField.Source is BitmapSource bitmapSource)
                    {
                        // Generate a unique filename for the image
                        string imageName = "UserImage.jpg";

                        // Save the image to isolated storage
                        using (var imageStream = new IsolatedStorageFileStream(imageName, FileMode.Create, isolatedStorage))
                        {
                            BitmapEncoder encoder = new JpegBitmapEncoder();
                            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                            encoder.Save(imageStream);
                        }

                        // Update the pdfImagePath with the isolated storage file name
                        pdfImagePath = imageName;
                        coImagePlaceholder.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        MessageBox.Show("Invalid image source.");
                        return;
                    }
                }
                finally
                {
                    // Ensure isolatedStorage is properly disposed
                    isolatedStorage?.Dispose();
                }

                // Update the UI elements with the saved data
                companyAddressValues[0] = tbAddL1.Text;
                companyAddressValues[1] = tbAddL2.Text;
                companyAddressValues[2] = tbAddL3.Text;
                companyAddressValues[3] = tbAddL4.Text;
                companyAddressValues[4] = tbTelNo.Text;

                coAddressL1P1.Text = tbAddL1.Text;
                coAddressL1P2.Text = tbAddL1.Text;
                coAddressL2P1.Text = tbAddL2.Text;
                coAddressL2P2.Text = tbAddL2.Text;
                coAddressL3P1.Text = tbAddL3.Text;
                coAddressL3P2.Text = tbAddL3.Text;
                coAddressL4P1.Text = tbAddL4.Text;
                coAddressL4P2.Text = tbAddL4.Text;
                coTelNumberP1.Text = tbTelNo.Text;
                coTelNumberP2.Text = tbTelNo.Text;

                MessageBox.Show("Saved");
                coDataSaved = true;
                CoDataBar.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                coDataSaved = false;
                MessageBox.Show($"An error occurred while saving user data: {ex.Message}");
            }
        }

        // Load user data from local storage
        private void LoadUserData()
        {
            // Open the isolated storage file if it exists
            using (var isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly())
            {
                //    isolatedStorage.DeleteFile("UserImage.jpg");//-------- Used for deleting Saved Company Data
                //    isolatedStorage.DeleteFile("UserData.txt");//-------- and so is this.
                if (isolatedStorage.FileExists("UserData.txt"))
                {
                    // Open the file and read its contents
                    using (var stream = new StreamReader(new IsolatedStorageFileStream("UserData.txt", FileMode.Open, isolatedStorage)))
                    {
                        string data = stream.ReadLine();

                        // Populate the textboxes with the loaded data
                        companyAddressValues = data.Split('|');
                        if (companyAddressValues.Length == 5)
                        {
                            // Replace the "&#44;" sequence with commas to restore the original values
                            for (int i = 0; i < companyAddressValues.Length; i++)
                            {
                                companyAddressValues[i] = companyAddressValues[i].Replace("&#44;", ",");
                            }
                            //MessageBox.Show("getting data from isolated storage");
                            tbAddL1.Text = companyAddressValues[0];
                            tbAddL2.Text = companyAddressValues[1];
                            tbAddL3.Text = companyAddressValues[2];
                            tbAddL4.Text = companyAddressValues[3];
                            tbTelNo.Text = companyAddressValues[4];
                        }
                        else
                        {
                            MessageBox.Show("Failed to get data from isolated storage");
                        }
                        coAddressL1P1.Text = tbAddL1.Text; coAddressL1P2.Text = tbAddL1.Text;
                        coAddressL2P1.Text = tbAddL2.Text; coAddressL2P2.Text = tbAddL2.Text;
                        coAddressL3P1.Text = tbAddL3.Text; coAddressL3P2.Text = tbAddL3.Text;
                        coAddressL4P1.Text = tbAddL4.Text; coAddressL4P2.Text = tbAddL4.Text;
                        coTelNumberP1.Text = tbTelNo.Text; coTelNumberP2.Text = tbTelNo.Text;

                        companyLogoP1.Source = imageUploadField.Source;
                        companyLogoP2.Source = imageUploadField.Source;
                    }
                    coDataSaved = true;
                }
                // Load the uploaded image from isolated storage
                if (isolatedStorage.FileExists("UserImage.jpg"))
                {
                    coImagePlaceholder.Visibility = Visibility.Hidden;
                    using (var imageStream = new IsolatedStorageFileStream("UserImage.jpg", FileMode.Open, isolatedStorage))
                    {
                        BitmapImage image = new BitmapImage();
                        image.BeginInit();
                        image.CacheOption = BitmapCacheOption.OnLoad;
                        image.StreamSource = imageStream;
                        image.EndInit();

                        imageUploadField.Source = image;

                        // Update the Image sources after the UI has rendered
                        Dispatcher.Invoke(() =>
                        {
                            companyLogoP1.Source = imageUploadField.Source;
                            companyLogoP2.Source = imageUploadField.Source;
                        });
                    }
                    coDataSaved = true;
                }
                else
                {
                    //MessageBox.Show("No image saved");
                }
            }
        }

        private void btnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png";

            if (openFileDialog.ShowDialog() == true)
            {
                string imagePath = openFileDialog.FileName;
                BitmapImage bitmap = new BitmapImage(new Uri(imagePath));

                // Update the imageUploadField.Source
                imageUploadField.Source = bitmap;

                // Update the companyLogoP1 and companyLogoP2 Sources
                companyLogoP1.Source = bitmap;
                companyLogoP2.Source = bitmap;
            }
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.InitialDirectory = @"C:\";
            dlg.Filter = "SJH Files (*.sjh)|*.sjh"; // Specify the crt file extension and description

            if (dlg.ShowDialog() == true)
            {
                string filePath = dlg.FileName;
                string folderPath = Path.GetDirectoryName(filePath);

                CertificateData loadedCertificateData = CertificateData.LoadDataFromFile(filePath);


                if (loadedCertificateData != null)
                {
                    imageSignatureData = loadedCertificateData.ImageSignatureData;
                    List<List<string>> formDataList = loadedCertificateData.FormDataList;
                    formDetails = loadedCertificateData.FormDetails;
                    // Update the original formDataArray with the loaded data
                    int rows = formDataArray.GetLength(0);
                    int columns = formDataArray.GetLength(1);

                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < columns; j++)
                        {
                            if (i < formDataList.Count && j < formDataList[i].Count)
                            {
                                formDataArray[i, j] = formDataList[i][j];
                            }
                            else
                            {
                                formDataArray[i, j] = string.Empty;
                            }
                        }
                    }
                    repopulateCert();
                }
                else
                {
                    Console.WriteLine("Failed to load the certificate data from the file.");
                }
            }
        }

        public void repopulateCert()
        {
            txtBoxReference.Text = formDetails[0];
            txtBoxCertNum.Text = formDetails[1];
            txtSiteAddress.Text = formDetails[2];
            comboEngineer.SelectedItem = formDetails[3];
            handleEngineerChanged();
            txtDate.Text = formDetails[8];

            for (int i = 0; i < 10; i++)
            {
                int index = i + 1;
                string sourceProtected = "tbsourceProtected" + index.ToString("D2");
                TextBox tbSourceProtected = FindName(sourceProtected) as TextBox;
                string circuitNo = "tbCircuitNo" + index.ToString("D2");
                TextBox tbCircuitNo = FindName(circuitNo) as TextBox;
                string man = "cbMan" + index.ToString("D2");
                ComboBox cbMan = FindName(man) as ComboBox;
                string prod = "cbProd" + index.ToString("D2");
                ComboBox cbProd = FindName(prod) as ComboBox;
                string l1pe = "tbl1pe" + index.ToString("D2");
                TextBox tbl1pe = FindName(l1pe) as TextBox;
                string l2pe = "tbl2pe" + index.ToString("D2");
                TextBox tbl2pe = FindName(l2pe) as TextBox;
                string l3pe = "tbl3pe" + index.ToString("D2");
                TextBox tbl3pe = FindName(l3pe) as TextBox;
                string npe = "tbnpe" + index.ToString("D2");
                TextBox tbnpe = FindName(npe) as TextBox;
                string l1n = "tbl1n" + index.ToString("D2");
                TextBox tbl1n = FindName(l1n) as TextBox;
                string l2n = "tbl2n" + index.ToString("D2");
                TextBox tbl2n = FindName(l2n) as TextBox;
                string l3n = "tbl3n" + index.ToString("D2");
                TextBox tbl3n = FindName(l3n) as TextBox;
                string zs = "tbzs" + index.ToString("D2");
                TextBox tbzs = FindName(zs) as TextBox;

                string location = "tbLocation" + index.ToString("D2");
                TextBox tbLocation = FindName(location) as TextBox;
                string BSEN = "tbBSEN" + index.ToString("D2");
                TextBox tbBSEN = FindName(BSEN) as TextBox;
                string isoType = "tbisotype" + index.ToString("D2");
                TextBox tbisotype = FindName(isoType) as TextBox;
                string rating = "tbRating" + index.ToString("D2");
                TextBox tbRating = FindName(rating) as TextBox;
                string SSC = "tbSSC" + index.ToString("D2");
                TextBox tbSSC = FindName(SSC) as TextBox;
                string RefMeth = "tbRefMeth" + index.ToString("D2");
                TextBox tbRefMeth = FindName(RefMeth) as TextBox;
                string live = "tbLive" + index.ToString("D2");
                TextBox tbLive = FindName(live) as TextBox;
                string CPC = "tbCPC" + index.ToString("D2");
                TextBox tbCPC = FindName(CPC) as TextBox;
                string r2 = "tbR2" + index.ToString("D2");
                TextBox tbR2 = FindName(r2) as TextBox;
                string irll = "tbIRLL" + index.ToString("D2");
                TextBox tbIRLL = FindName(irll) as TextBox;
                string irlpe = "tbIRLPE" + index.ToString("D2");
                TextBox tbIRLPE = FindName(irlpe) as TextBox;
                string P = "cbP" + index.ToString("D2");
                CheckBox cbP = FindName(P) as CheckBox;
                string comPass = "PassCheckBox" + index.ToString("D2");
                CheckBox PassCheckBox = FindName(comPass) as CheckBox;
                string comFail = "Group1CheckBoxes" + index.ToString("D2");
                StackPanel Group1CheckBoxes = FindName(comFail) as StackPanel;

                tbSourceProtected.Text = formDataArray[i, 0];
                tbCircuitNo.Text = formDataArray[i, 1];
                cbMan.SelectedItem = formDataArray[i, 2];
                cbProd.SelectedItem = formDataArray[i, 3];
                tbl1pe.Text = formDataArray[i, 6];
                tbl2pe.Text = formDataArray[i, 8];
                tbl3pe.Text = formDataArray[i, 10];
                tbnpe.Text = formDataArray[i, 12];
                tbl1n.Text = formDataArray[i, 14];
                tbl2n.Text = formDataArray[i, 16];
                tbl3n.Text = formDataArray[i, 18];
                tbzs.Text = formDataArray[i, 19];

                tbLocation.Text = formDataArray[i, 20];
                tbBSEN.Text = formDataArray[i, 21];
                tbisotype.Text = formDataArray[i, 22];
                tbRating.Text = formDataArray[i, 23];
                tbSSC.Text = formDataArray[i, 24];
                tbRefMeth.Text = formDataArray[i, 25];
                tbLive.Text = formDataArray[i, 26];
                tbCPC.Text = formDataArray[i, 27];
                tbR2.Text = formDataArray[i, 28];
                tbIRLL.Text = formDataArray[i, 29];
                tbIRLPE.Text = formDataArray[i, 30];
                string value = formDataArray[i, 31];
                if (formDataArray[i, 0] != null)
                {
                    if (bool.TryParse(value, out bool isChecked))
                    {
                        cbP.IsChecked = isChecked;
                    }
                    else
                    {
                        cbP.IsChecked = null; // Set to indeterminate state
                    }
                }
                if (formDataArray[i, 32] != null)
                {
                    //MessageBox.Show(formDataArray[i, 32]);
                    if (formDataArray[i, 32] == "No remedial action required")
                    {
                        PassCheckBox.IsChecked = true;
                    }
                    if (formDataArray[i, 32].Contains("C3 - code 1"))
                    {
                        CheckBox firstCheckBox = Group1CheckBoxes.Children[0] as CheckBox;
                        if (firstCheckBox != null)
                        {
                            firstCheckBox.IsChecked = true;
                        }

                    }

                    if (formDataArray[i, 32].Contains("C2 - code(s)"))
                    {
                        // Check the relevant checkboxes
                        foreach (CheckBox checkBox in Group1CheckBoxes.Children.OfType<CheckBox>())
                        {
                            if (checkBox.Name != "NumberedCheckBox" + index.ToString("D2"))
                            {
                                string checkBoxContent = checkBox.Content.ToString();
                                if (formDataArray[i, 32].Contains("C2 - code(s) " + checkBoxContent))
                                {
                                    checkBox.IsChecked = true;
                                }
                            }
                        }
                    }

                }
            }

            MessageBox.Show("Data Loaded.");
        }

        private void infoButton_Click(object sender, RoutedEventArgs e)
        {
            FailCodeGuide failCodeGuide = new FailCodeGuide();
            failCodeGuide.Topmost = true;
            failCodeGuide.Show();
        }

        private void logoButton_Click(object sender, RoutedEventArgs e)
        {
            LogoGuide logoGuide = new LogoGuide();
            logoGuide.Topmost = true;
            logoGuide.Show();
        }

        private void updateBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateManager updateManager = new UpdateManager();
            _ = updateManager.CheckForUpdates();
        }

        private void signatureButton_Click(object sender, RoutedEventArgs e)
        {
            SignatureGuide signatureGuide = new SignatureGuide();
            signatureGuide.Topmost = true;
            signatureGuide.Show();
        }

        //------------ Zoom control -----------

        private void MainWindow_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var transforms = GetCurrentTransforms();
            var viewBox = GetCurrentViewBox();

            // If no transforms or ViewBox, do nothing
            if (!transforms.HasValue || viewBox == null) return;

            var (scaleTransform, translateTransform) = transforms.Value;

            double newScaleX = scaleTransform.ScaleX;
            double newScaleY = scaleTransform.ScaleY;

            if (e.Delta > 0)  // zoom in
            {
                newScaleX *= 1.1;
                newScaleY *= 1.1;
            }
            else  // zoom out
            {
                newScaleX /= 1.1;
                newScaleY /= 1.1;
            }

            // clamp the scale between 1 (100%) and 3 (300%)
            newScaleX = Math.Max(1, Math.Min(newScaleX, 3));
            newScaleY = Math.Max(1, Math.Min(newScaleY, 3));

            // Calculate the mouse position difference
            var position = e.GetPosition(viewBox);
            var diffX = position.X - translateTransform.X;
            var diffY = position.Y - translateTransform.Y;

            // Calculate new translate transform (zoom towards mouse position)
            var newTranslateX = position.X - diffX * newScaleX / scaleTransform.ScaleX;
            var newTranslateY = position.Y - diffY * newScaleY / scaleTransform.ScaleY;

            // Check if the content is going out of view and adjust the position
            newTranslateX = Math.Min(Math.Max(newTranslateX, -viewBox.ActualWidth * (newScaleX - 1)), 0);
            newTranslateY = Math.Min(Math.Max(newTranslateY, -viewBox.ActualHeight * (newScaleY - 1)), 0);

            // Only apply the new scale and translate if they have changed
            if (newScaleX != scaleTransform.ScaleX || newScaleY != scaleTransform.ScaleY ||
                newTranslateX != translateTransform.X || newTranslateY != translateTransform.Y)
            {
                scaleTransform.ScaleX = newScaleX;
                scaleTransform.ScaleY = newScaleY;
                translateTransform.X = newTranslateX;
                translateTransform.Y = newTranslateY;
            }
        }


        private Point? dragStart = null;

        private void Viewbox_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var transforms = GetCurrentTransforms();
            var viewBox = GetCurrentViewBox();

            // If no transforms or ViewBox, do nothing
            if (!transforms.HasValue || viewBox == null) return;

            var (scaleTransform, translateTransform) = transforms.Value;
            // Only start panning on left mouse button
            if (e.ChangedButton == MouseButton.Left)
            {
                // Change the cursor to a hand to indicate that panning has started
                this.Cursor = Cursors.Hand;

                // Remember the point where the mouse down event occurred
                dragStart = e.GetPosition(viewBox);
                e.Handled = true;
            }
        }

        private void Viewbox_MouseMove(object sender, MouseEventArgs e)
        {
            var transforms = GetCurrentTransforms();
            var viewBox = GetCurrentViewBox();

            // If no transforms or ViewBox, do nothing
            if (!transforms.HasValue || viewBox == null) return;

            var (scaleTransform, translateTransform) = transforms.Value;
            // If panning is in progress and the mouse is still down
            if (dragStart != null && e.LeftButton == MouseButtonState.Pressed && scaleTransform.ScaleX > 1)
            {
                // Get the position of the mouse
                var currentPosition = e.GetPosition(viewBox);

                // Calculate the distance the mouse has moved
                var offsetX = (currentPosition.X - dragStart.Value.X) * 1.5;
                var offsetY = (currentPosition.Y - dragStart.Value.Y) * 1.5;

                // Calculate the new position
                var newX = translateTransform.X + offsetX;
                var newY = translateTransform.Y + offsetY;

                // Prevent dragging off screen
                newX = Math.Min(Math.Max(newX, -viewBox.ActualWidth * (scaleTransform.ScaleX - 1)), 0);
                newY = Math.Min(Math.Max(newY, -viewBox.ActualHeight * (scaleTransform.ScaleY - 1)), 0);

                // Update the TranslateTransform
                translateTransform.X = newX;
                translateTransform.Y = newY;

                // Update the start position
                dragStart = currentPosition;

                e.Handled = true;
            }
        }

        private void Viewbox_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var transforms = GetCurrentTransforms();
            var viewBox = GetCurrentViewBox();

            // If no transforms or ViewBox, do nothing
            if (!transforms.HasValue || viewBox == null) return;

            var (scaleTransform, translateTransform) = transforms.Value;
            // Change the cursor back to the arrow when panning is finished
            this.Cursor = Cursors.Arrow;

            // Stop panning
            dragStart = null;
            e.Handled = true;
        }

        private (ScaleTransform, TranslateTransform)? GetCurrentTransforms()
        {
            if (tabControl.SelectedItem == tabPage1)
            {
                return (scaleTransformP1, translateTransformP1);
            }
            else if (tabControl.SelectedItem == tabPage2)
            {
                return (scaleTransformP2, translateTransformP2);
            }

            return null; // No transform for other tabs
        }

        private Viewbox GetCurrentViewBox()
        {
            if (tabControl.SelectedItem == tabPage1)
            {
                return viewBoxPage1;
            }
            else if (tabControl.SelectedItem == tabPage2)
            {
                return viewBoxPage2;
            }

            return null; // No ViewBox for other tabs
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            bool pdbSelected = false;
            bool edbSelected = false;

            DatabaseOptionsDialogue dialog = new DatabaseOptionsDialogue();

            if (dialog.ShowDialog() == true)
            {
                pdbSelected = dialog.OptionPDB;
                edbSelected = dialog.OptionEDB;

                // Use the selected options as needed
                MessageBox.Show($"Product Database selected: {pdbSelected}, Engineers Database selected: {edbSelected}");
            }
            using (System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                try
                {
                    folderDialog.Description = "Select the directory to save the backup.";
                    if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string backupPath = folderDialog.SelectedPath;
                        if (pdbSelected)
                        {
                            string backupFilePath = Path.Combine(backupPath, "ProductDataBackup.bak");
                            string connString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ProductData.mdf;Integrated Security=True";
                            string mdfFilePath = Path.Combine(backupPath, $"ProductData.mdf");
                            string ldfFilePath = Path.Combine(backupPath, $"ProductData_log.ldf");
                            LogErrorToFile("Calling CreateBAKFile for ProductData");
                            LogErrorToFile("backupFilePath = " + backupFilePath);
                            CreateBAKFile(backupFilePath, connString, mdfFilePath,ldfFilePath);

                        }

                        if (edbSelected)
                        {
                            string backupFilePath = Path.Combine(backupPath, "EngineersDataBackup.bak");
                            string connString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\EngineersData.mdf;Integrated Security=True";
                            string mdfFilePath = Path.Combine(backupPath, $"EngineersData.mdf");
                            string ldfFilePath = Path.Combine(backupPath, $"EngineersData_log.ldf");
                            LogErrorToFile("Calling CreateBAKFile for EngineersData");
                            LogErrorToFile("backupFilePath = " + backupFilePath);
                            CreateBAKFile(backupFilePath, connString, mdfFilePath, ldfFilePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogErrorToFile(ex.Message);
                    MessageBox.Show($"See log file. An error occurred: {ex.Message}");
                }
                }
        }

        private void CreateBAKFile(string backupFilePath, string connString, string mdfFilePath, string ldfFilePath)
        {
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();

                // Database name to restore
                string dbNameToRestore = Path.GetFileNameWithoutExtension(mdfFilePath);

                // Detach or Drop the database if it exists
                string detachOrDropDbSQL = $"IF EXISTS (SELECT * FROM sys.databases WHERE name = '{dbNameToRestore}') " +
                                           $"BEGIN " +
                                           $"ALTER DATABASE [{dbNameToRestore}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; " +
                                           $"DROP DATABASE [{dbNameToRestore}]; " +
                                           $"END";
                using (SqlCommand cmd = new SqlCommand(detachOrDropDbSQL, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // Check current database name
                string dbName;
                using (SqlCommand cmd = new SqlCommand("SELECT DB_NAME() AS [Current Database]", conn))
                {
                    dbName = cmd.ExecuteScalar().ToString();
                    MessageBox.Show($"Current database is {dbName}");
                }

                // Perform the backup
                string backupSQL = $"BACKUP DATABASE [{dbName}] TO DISK = @backupFilePath";
                using (SqlCommand cmd = new SqlCommand(backupSQL, conn))
                {
                    cmd.Parameters.AddWithValue("@backupFilePath", backupFilePath);
                    cmd.ExecuteNonQuery();
                }
                LogErrorToFile("BAK file created");

                // Get logical file names from the backup
                string logicalMdfName = "", logicalLdfName = "";
                string restoreFileListSQL = $"RESTORE FILELISTONLY FROM DISK = '{backupFilePath}'";
                using (SqlCommand cmd = new SqlCommand(restoreFileListSQL, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string type = reader["Type"].ToString();
                        if (type == "D") // Database
                            logicalMdfName = reader["LogicalName"].ToString();
                        else if (type == "L") // Log
                            logicalLdfName = reader["LogicalName"].ToString();
                    }
                }
                LogErrorToFile("BAK file contents = " + logicalMdfName + " and " + logicalLdfName);

                // Perform the restore
                string restoreSQL = $"RESTORE DATABASE [{dbNameToRestore}] " +
                                    $"FROM DISK = '{backupFilePath}' " +
                                    $"WITH MOVE '{logicalMdfName}' TO '{mdfFilePath}', " +
                                    $"MOVE '{logicalLdfName}' TO '{ldfFilePath}', " +
                                    "REPLACE";
                using (SqlCommand cmd = new SqlCommand(restoreSQL, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }


        private void btnRestore_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Do you want to start the restore process the next time you start the application? Selecting 'Yes' will set the application to restore mode on next startup.",
        "Restore on Next Startup", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                SetRestoreFlag(true);
                MessageBox.Show("The application will start in restore mode on the next startup.", "Restore Mode Set", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SetRestoreFlag(bool shouldRestore)
        {
            using (IsolatedStorageFile isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly())
            {
                using (IsolatedStorageFileStream stream = new IsolatedStorageFileStream("RestoreFlag.txt", FileMode.Create, isolatedStorage))
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.WriteLine(shouldRestore);
                }
            }
        }

        private bool CheckForRestoreFlag()
        {
            using (IsolatedStorageFile isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly())
            {
                if (isolatedStorage.FileExists("RestoreFlag.txt"))
                {
                    using (IsolatedStorageFileStream stream = new IsolatedStorageFileStream("RestoreFlag.txt", FileMode.Open, isolatedStorage))
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string flag = reader.ReadLine();
                        return bool.TryParse(flag, out bool shouldRestore) && shouldRestore;
                    }
                }
            }
            return false;
        }

        private void restoreFunction()
        {
            //    OpenFileDialog openFileDialog = new OpenFileDialog
            //    {
            //        Filter = "Backup Files|*.bak",
            //        Title = "Select Backup File"
            //    };

            string dataDirectory = AppDomain.CurrentDomain.GetData("DataDirectory") as string;
            if (string.IsNullOrEmpty(dataDirectory))
            {
                dataDirectory = AppDomain.CurrentDomain.BaseDirectory;
            }
            string mdfPath = Path.Combine(dataDirectory, "ProductData.mdf");
            string ldfPath = Path.Combine(dataDirectory, "ProductData_log.ldf");
            

            try
            {
                if (File.Exists(mdfPath))
                {
                    File.Delete(mdfPath);
                }

                if (File.Exists(ldfPath))
                {
                    File.Delete(ldfPath);
                }

                MessageBox.Show("Database files deleted successfully from: " + mdfPath);
                SetRestoreFlag(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while deleting database files: " + ex.Message);
            }
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Backup Files|*.*",
                Title = "Select Backup File"
            };

            //    if (openFileDialog.ShowDialog() == true)
            //    {
            //        string cnst = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ProductData.mdf;Integrated Security=True";
            //        RestoreDatabaseFromBackup(openFileDialog.FileName, "ProductData", cnst);

            //    }
        }





    private void LogErrorToFile(string errorMessage)
        {
            string logFilePath = @"C:\Users\alanh\Desktop\Boobies\errorLog.txt";  // Change this to the path where you want to save the log file
            using (StreamWriter sw = File.AppendText(logFilePath))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + errorMessage);
            }
        }
    }
}





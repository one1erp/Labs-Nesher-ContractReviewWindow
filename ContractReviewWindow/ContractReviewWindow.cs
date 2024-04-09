using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Objects.DataClasses;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Windows.Forms;
using Common;
using Validators;

using DAL;
using LSExtensionWindowLib;
using LSSERVICEPROVIDERLib;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using One1.Controls;
using Telerik.WinControls.UI;
using Application = Microsoft.Office.Interop.Word.Application;
using Exception = System.Exception;
using Range = Microsoft.Office.Interop.Word.Range;
using Row = Microsoft.Office.Interop.Word.Row;


namespace ContractReviewWindow
{
    [ComVisible(true)]
    [ProgId("ContractReviewWindow.WindowExtension")]
    public partial class ContractReviewControl : UserControl, IExtensionWindow
    {
        #region Ctor

        public ContractReviewControl()
        {

            InitializeComponent();
            BackColor = Color.FromName("Control");
  
        }

        #endregion

        #region private members

        private string ClientPath;
        private INautilusDBConnection NtlsCon;
        private IExtensionWindowSite2 NtlsSite;
        private INautilusProcessXML ProcessXML;
        private Client _currentClient;
        private List<TestTemplateEx> _testTemplateList;
        private IDataLayer dal;
        private Document doc;
        private Contract lastContract;
        private RadGridView grdAllTests;
        private PhraseHeader docTEmplates;
        private string _generalContractLocation;

        #endregion

        #region Implementation of IExtensionWindow

        public bool CloseQuery()
        {
            return true;
        }

        public void Internationalise()
        {
        }

        public void SetSite(object site)
        {
            NtlsSite = (IExtensionWindowSite2)site;
            NtlsSite.SetWindowInternalName("סקר חוזה");
            NtlsSite.SetWindowRegistryName("סקר חוזה");
            NtlsSite.SetWindowTitle("סקר חוזה");
        }


        public void PreDisplay()
        {
            Utils.CreateConstring(NtlsCon);
            Connect();
        }

        public WindowButtonsType GetButtons()
        {
            return WindowButtonsType.windowButtonsNone;
        }

        public bool SaveData()
        {
            return false; //???
        }

        public void SetServiceProvider(object serviceProvider)
        {
            var sp = serviceProvider as NautilusServiceProvider;
            ProcessXML = Utils.GetXmlProcessor(sp);
            NtlsCon = Utils.GetNtlsCon(sp);
        }

        public void SetParameters(string parameters)
        {
        }

        public void Setup()
        {
        }

        public WindowRefreshType DataChange()
        {
            return WindowRefreshType.windowRefreshNone;
        }

        public WindowRefreshType ViewRefresh()
        {
            return WindowRefreshType.windowRefreshNone;
        }

        public void refresh()
        {
        }

        public void SaveSettings(int hKey)
        {
        }

        public void RestoreSettings(int hKey)
        {
        }

        public void Close()
        {
        }

        #endregion

        #region private methods

        private void Connect()
        {
            dal = new DataLayer();
            dal.Connect();
            GetTestTemplate();
            GetCompliteonPhrase();
            GetClients();
            LoadDropDownTemplates();
            GetFolderLocations();
            dropDownClients.SelectedIndexChanged += DropDownClients_SelectedValueChanged;
        }

        private void GetCompliteonPhrase()
        {
            CompletionPhrase = dal.GetPhraseByName("Completion");

        }
      
        private PhraseHeader CompletionPhrase;
        private string clientDocumentsPath;

        private void GetFolderLocations()
        {
            try
            {
                var phraseH = dal.GetPhraseByName("Location folders");
                var pe = phraseH.PhraseEntries.Where(p => p.PhraseDescription == "Client documents").FirstOrDefault();
                clientDocumentsPath = pe.PhraseName;
            }
            catch (Exception e)
            {

                CustomMessageBox.Show("טעות בהגדרות המערכת, אנא פנה לתמיכה");
                Logger.WriteLogFile(e);
            }
        }

        private void GetTestTemplate()
        {
            PopulateTabs();
        }

        private void GetClients()
        {
            List<Client> _clients = dal.GetClients();

            dropDownClients.DisplayMember = "Name";

            dropDownClients.DataSource = _clients;

            dropDownClients.SelectedIndex = -1;
        }

        private void GetContractByClient()
        {
            lastContract = (from item in _currentClient.Contracts
                                                        .OrderByDescending(x => x.ConfirmDate)
                            select item).FirstOrDefault();


            if (lastContract != null)
            {
                radDateTimePicker1.Visible = true;
                lblNoContract.Visible = false;
                if (lastContract.ConfirmDate != null) radDateTimePicker1.Value = (DateTime)lastContract.ConfirmDate;
                SetGridSource(lastContract);
                textBoxRemarks.Text = lastContract.U_REMARKS;


            }


            else
            {
                lastContract = new Contract();

                radDateTimePicker1.Visible = false;
                lblNoContract.Visible = true;

                ContractGrid.DataSource = null;
                btnOpenDoc.Enabled = false;
                btnSendEmail.Enabled = false;
                textBoxRemarks.Text = string.Empty;

                // SetEnableButtons(false);
                btnDiscount.Enabled = false;
            }
            enableButtonsFromContractGrid();
        }

        private void SetGridSource(Contract contract)
        {
            ContractGrid.BeginUpdate();
            ContractGrid.DataSource = null;
            ContractGrid.DataSource = contract.ContractDatas.Where(x => x.EntityState != EntityState.Deleted);
            ContractGrid.EndUpdate();
        }

        private void PopulateTabs()
        {
            //get Test templates
            _testTemplateList = dal.GetTestTemplatesForPriceList();

            var labinfos = dal.GetLabs();

            //Add tab for all tests
            AddTab("כל הבדיקות", "ALL");
            foreach (var labInfo in labinfos)
            {
                AddTab(labInfo.LabHebrewName, labInfo.LabLetter);
            }

            foreach (var tab in radPageLabs.Pages)
            {
                //Create a RadGridView for each TAB
                CreateGridView(tab);

                grdAllTests = (RadGridView)tab.Controls[0];

                //Add coulmns
                AddColumns(grdAllTests);

                DesignColumns(grdAllTests);

                //get lab of tab
                string labTab = tab.Tag.ToString();


                if (labTab == "ALL") //כל הבדיקות
                {
                    grdAllTests.DataSource = _testTemplateList;
                }
                else
                {
                    var newList = new List<TestTemplateEx>();

                    foreach (TestTemplateEx item in _testTemplateList)
                    {
                        if (item.RelevantLabs == null) continue;
                        string[] labs = item.RelevantLabs.Split(',');

                        bool containsThisLab = labs.Contains(labTab);
                        if (containsThisLab)
                        {
                            newList.Add(item);
                        }
                    }

                    //Set data source
                    grdAllTests.DataSource = newList;
                }
            }

        }

        private void AddTab(string text, string tag)
        {
            var tabAll = new RadPageViewPage();
            tabAll.Text = text;
            tabAll.Tag = tag;
            radPageLabs.Pages.Add(tabAll);
        }



        private void CreateGridView(RadPageViewPage tab)
        {
            var rgv = new RadGridView();
            rgv.EnableAlternatingRowColor = true;
            rgv.EnableFiltering = true;
            rgv.ShowFilteringRow = true;
            rgv.ShowRowErrors = true;
            rgv.MultiSelect = true;
            rgv.Dock = DockStyle.Fill;
            rgv.AllowAddNewRow = false;
            rgv.RightToLeft = RightToLeft.Yes;
            rgv.ShowRowHeaderColumn = true;
            rgv.AllowDeleteRow = false;
            rgv.Width = 800;
            rgv.Height = 800;
            rgv.AllowDragToGroup = false;
            rgv.ReadOnly = true;
            rgv.AutoGenerateColumns = true;
            rgv.TableElement.Padding = new Padding(0);
            rgv.TableElement.DrawBorder = true;
            rgv.TableElement.CellSpacing = -1;
            rgv.TableElement.Text = "";
            rgv.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.Fill;
            tab.Controls.Add(rgv);
        }

        private void AddColumns(RadGridView radDataGrid)
        {
            if (radDataGrid.ColumnCount == 0)
            {
                radDataGrid.Columns.Add(new GridViewTextBoxColumn("שם הבדיקה", "Name"));
                radDataGrid.Columns.Add(new GridViewDecimalColumn("מחיר", "PRICE"));
            }
        }

        private void DesignColumns(RadGridView radDataGrid)
        {
            foreach (GridViewDataColumn item in radDataGrid.Columns)
            {
                item.Width = 100;
                item.TextAlignment = ContentAlignment.MiddleCenter;
            }
        }


        private void LoadDropDownTemplates()
        {
            docTEmplates = dal.GetPhraseByName("Contract documents templates");
            ddlFileTemplate.DisplayMember = "PhraseName";
            ddlFileTemplate.DataSource = docTEmplates.PhraseEntries;
        }

        private void ClearClientData()
        {
            txtDiscount.Text = "";
            txtContactMan.Text = "";
            txtPhone.Text = "";
            linkEmail.Text = "";
            txtClientCode.Text = "";
            txtAddress.Text = "";
            textBoxRemarks.Text = "";
            txtPath.Text = "";
            ContractGrid.DataSource = null;

            SetEnableButtons(false);
            ClientPath = string.Empty;
        }

        private void SetEnableButtons(bool flag)
        {
            //  Utils.UiHelperMethods.SetEnabled(this, flag, "btnOpenDoc",
            //                                              "btnRemoveFromContract"

            //, "btnAddToContract"
            //, "btnChooseAllContractT"
            //, "btnChooseAlltetstT"
            //, "btnClearAlltetstT"
            //, "btnClearAllContractT"
            //, "btnDiscount"
            //, "btnSave"
            //, "btnSendEmail")
            //  ;

            btnOpenDoc.Enabled = flag;
            btnRemoveFromContract.Enabled = flag;
            btnAddToContract.Enabled = flag;
            btnChooseAllContractT.Enabled = flag;
            btnChooseAlltetstT.Enabled = flag;
            btnClearAlltetstT.Enabled = flag;
            btnClearAllContractT.Enabled = flag;
            btnDiscount.Enabled = flag;
            btnSave.Enabled = flag;
            btnSendEmail.Enabled = flag;


        }


        private void CalculatePercentage(ContractData cd, decimal price)
        {
            ContractGrid.BeginUpdate();

            if (cd.TestTemplateEx.Price != null)
            {
                decimal dis = ((decimal)cd.TestTemplateEx.Price * price / 100);
                cd.FinalPrice = cd.TestTemplateEx.Price - dis;
            }

            ContractGrid.EndUpdate();
        }

        private void CloseWindow()
        {
            dal.Close();
            NtlsSite.CloseWindow();
        }

        #endregion

        #region Events

        private void DropDownClients_SelectedValueChanged(object sender, EventArgs e)
        {
            if (dropDownClients.SelectedIndex > -1)
            {
                if (dal.HasChanges())
                {
                    DialogResult dr =
                        MessageBox.Show(string.Format("נעשו שינויים בלקוח ,{0} \n  ?האם ברצונך לשמור את השינויים",
                                                      _currentClient.Name), "Nautilus", MessageBoxButtons.YesNoCancel);
                    if (dr == DialogResult.Yes)
                    {
                        btnSave_Click(null, null);
                    }
                    else if (dr == DialogResult.No)
                    {
                        if (lastContract != null) dal.CancelLastContract(lastContract);
                        //Refresh 
                        dal.RefreshContract(lastContract);
                    }
                }

                ClearClientData();


                //Get selected client
                var client = (Client)dropDownClients.SelectedItem.DataBoundItem;
                _currentClient = client;
                txtDiscount.Text = client.Discount.ToString();
                txtPath.Text = client.ContractFileName;

                txtPath.Enabled = string.IsNullOrEmpty(client.ContractFileName);

                Address address = dal.GetAddresses("CLIENT", client.ClientId).FirstOrDefault(ad => ad.AddressType == "C");
                if (address != null)
                {
                    txtDiscount.Text = client.Discount.ToString();
                    txtContactMan.Text = address.ContactMan;
                    txtAddress.Text = address.FullAddress;
                    txtPhone.Text = address.Phone;
                    linkEmail.Text = address.Email;
                    txtClientCode.Text = client.ClientCode;
                }
                SetEnableButtons(true);
                GetContractByClient();
                Focus();
            }
            else
                ClearClientData();
        }

        private void btnAddToContract_Click(object sender, EventArgs e)
        {
            //Get selected tab
            RadPageViewPage selectedTab = radPageLabs.SelectedPage;
            //Get gridview from tab
            var radDataGrid = (RadGridView)selectedTab.Controls[0];

            //Get selected items to move to contract
            GridViewSelectedRowsCollection selected = radDataGrid.SelectedRows;
            if (selected.Count > 0)
            {
                btnDiscount.Enabled = true;


                foreach (GridViewRowInfo item in selected)
                {
                    var tt = (TestTemplateEx)item.DataBoundItem;


                    if (
                        lastContract != null && lastContract.ContractDatas.Count(
                            cd => cd != null && cd.TestTemplateEx != null && cd.TestTemplateExId == tt.TestTemplateExId) < 1)
                    {
                        lastContract.ContractDatas.Add(new ContractData
                                                           {
                                                               TestTemplateEx = tt,
                                                               Remarks = "",
                                                               FinalPrice = tt.Price,
                                                           });
                    }
                    else
                    {
                        //  MessageBox.Show("הבדיקה כבר קיימת.", "Nautilus");
                    }
                }
                SetGridSource(lastContract);
            }
            enableButtonsFromContractGrid();
        }

        private void enableButtonsFromContractGrid()
        {
            bool b = ContractGrid.RowCount > 0;
            btnRemoveFromContract.Enabled = b;
            btnChooseAllContractT.Enabled = b;
            btnClearAllContractT.Enabled = b;
        }

        private void btnRemoveFromContract_Click(object sender, EventArgs e)
        {
            GridViewSelectedRowsCollection selected = ContractGrid.SelectedRows;

            if (selected.Count > 0)
            {
                string str = string.Format("האם אתה בטוח שברצונך להוריד " + selected.Count + "? בדיקות מהחוזה");
                DialogResult dr = MessageBox.Show(str, "הורדת בדיקה",
                                                  MessageBoxButtons.YesNoCancel,
                                                  MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    foreach (
                        ContractData contractData in selected.ToList().Select(item => (ContractData)item.DataBoundItem)
                        )
                    {
                        dal.Delete(contractData);
                    }

                    SetGridSource(lastContract);
                }
            }
            enableButtonsFromContractGrid();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (dal.HasChanges())
            {
                DialogResult dr = MessageBox.Show(string.Format(" ? האם ברצונך לשמור את השינויים שנעשו"), "Nautilus",
                                                  MessageBoxButtons.YesNoCancel);
                if (dr == DialogResult.Yes)
                {
                    btnSave_Click(null, null);
                    CloseWindow();
                }
                else
                {
                    if (dr == DialogResult.No)
                        CloseWindow();
                }
            }
            else
            {
                DialogResult dr = MessageBox.Show("האם אתה בטוח שברצונך לצאת מהמהסך ?", "Nautilus",
                                                  MessageBoxButtons.YesNoCancel);
                if (dr == DialogResult.Yes)
                {
                    CloseWindow();
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPath.Text))
            {
                CustomMessageBox.Show("אנא הגדר תיקיית יעד עבור הלקוח");
                txtPath.Focus();

                return;
            }

            btnSave.Enabled = false;
            if (dal.HasChanges())
            {

                //Clone data
                var newContract = new Contract();                
                newContract.VERSION_STATUS = "A";               
                newContract.ConfirmDate = DateTime.Now;
                newContract.ContractId = dal.GetNewContractID();
                newContract.Name = newContract.ContractId.ToString();
                newContract.U_REMARKS = textBoxRemarks.Text;
                foreach (ContractData contractData in lastContract.ContractDatas.ToList())
                {
                    
                    var newContractData = new ContractData
                                              {
                                                  TestTemplateEx = contractData.TestTemplateEx,
                                                  Remarks = contractData.Remarks,
                                                  FinalPrice = contractData.FinalPrice,
                                                  ContractDataId = dal.GetNewContractDataID()
                                                 
                                              };
                    newContract.ContractDatas.Add(newContractData);
                }
                _currentClient.Contracts.Add(newContract);

                //Cancel changes from last contract
                dal.CancelLastContract(lastContract);
                if (string.IsNullOrEmpty(_currentClient.ContractFileName))
                    _currentClient.ContractFileName = txtPath.Text.MakeSafeFilename('_').Trim(' ');


                //Save Discount 
                decimal discount;
                var b = decimal.TryParse(txtDiscount.Text, out discount);

                _currentClient.Discount = discount;

                dal.SaveChanges();

                //Refresh grid 
                GetContractByClient();

                txtPath.Enabled = false;


                CreateSeker(newContract);
                //ClearClientData();


                btnOpenDoc.Enabled = true;
            }
            else
            {
                CustomMessageBox.Show("לא נעשו שינויים", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            btnSave.Enabled = true;

        }



        private void btnChooseAllContractT_Click(object sender, EventArgs e)
        {
            ContractGrid.SelectAll();
        }

        private void btnClearAllContractT_Click(object sender, EventArgs e)
        {
            ContractGrid.ClearSelection();
        }

        private void btnChooseAlltetstT_Click(object sender, EventArgs e)
        {
            RadPageViewPage selectedTab = radPageLabs.SelectedPage;
            var radDataGrid = (RadGridView)selectedTab.Controls[0];
            radDataGrid.SelectAll();
        }

        private void btnClearAlltetstT_Click(object sender, EventArgs e)
        {
            RadPageViewPage selectedTab = radPageLabs.SelectedPage;
            var radDataGrid = (RadGridView)selectedTab.Controls[0];
            radDataGrid.ClearSelection();
        }

        private void btnDiscount_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("  האם אתה בטוח שברצונך להחיל הנחת לקוח על כל הבדיקות?", "Nautilus",
                                              MessageBoxButtons.YesNoCancel,
                                              MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (dr == DialogResult.Yes)
            {
                decimal price;
                bool hasDiscount = decimal.TryParse(txtDiscount.Text, out price) && price > 0;
                if (hasDiscount)
                {
                    GridViewRowCollection selected = ContractGrid.Rows;

                    foreach (GridViewRowInfo item in selected.ToList())
                    {
                        var cd = (ContractData)item.DataBoundItem;


                        CalculatePercentage(cd, price);
                    }
                }
            }
        }


        private void radGridView1_CellValueChanged_1(object sender, GridViewCellEventArgs e)
        {
            btnRemoveFromContract.Enabled = ContractGrid.RowCount > 0;
        }

        private void radGridView1_CellValueChanged_1(object sender, EventArgs e)
        {
            object dataBoundItem = (((GridViewCellEventArgsBase)(e)).Row).DataBoundItem;
            var cd = (ContractData)dataBoundItem;
            decimal price;
            bool hasDiscount = decimal.TryParse(txtDiscount.Text, out price) && price > 0;
            if (hasDiscount)
            {
                CalculatePercentage(cd, price);
            }
        }

        private void radPageLabs_SelectedPageChanged(object sender, EventArgs e)
        {
            //Set auto complete to text box search
            //            AddAutoCompleteTextBox();
        }

        private void radGridView1_CommandCellClick(object sender, EventArgs e)
        {
            object dataBoundItem = (((GridViewCellEventArgsBase)(e)).Row).DataBoundItem;
            var cd = (ContractData)dataBoundItem;
            decimal price;
            bool hasDiscount = decimal.TryParse(txtDiscount.Text, out price) && price > 0;
            if (hasDiscount)
            {
                CalculatePercentage(cd, price);
            }
        }

        private void ContractReviewControl_Resize(object sender, EventArgs e)
        {
            lblHeader.Location = new System.Drawing.Point(Width / 2 - lblHeader.Width / 2, lblHeader.Location.Y);
        }

        #endregion

        #region Document

        private void CreateSeker(Contract newContract)
        {
            if (!string.IsNullOrEmpty(_currentClient.ContractFileName))
            {


                //Get selected template
                //Old
                //   var pe = (PhraseEntry)ddlFileTemplate.SelectedValue;
                //    string templatefileName = pe.PhraseDescription;




                var templatefileName = docTEmplates.PhraseEntries.FirstOrDefault().PhraseDescription;

                string savedFileName = CreateSavedFileName(templatefileName);


                Application wordApp = null;
                Document document = null;
                Document newDocument = null;

                try
                {
                    wordApp = new Application();
                    wordApp.Visible = false;

                    document = wordApp.Documents.Open(templatefileName);


                    newDocument = CopyToNewDocument(document);

                    EntityCollection<ContractData> itemsToUpdate = newContract.ContractDatas;
                    //   Change it to new doc
                    foreach (var contractData in itemsToUpdate)
                    {
                        Row row = newDocument.Tables[1].Rows.Add();
                        row.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                        row.Cells[1].Range.Text = contractData.TestTemplateEx.Name;
                        row.Cells[2].Range.Text = contractData.TestTemplateEx.Standard;
                        var phraseEntry =
                            CompletionPhrase.PhraseEntries.Where(
                                x => x.PhraseName == contractData.TestTemplateEx.Completion).FirstOrDefault();
                        if (phraseEntry != null)
                            row.Cells[3].Range.Text = phraseEntry.PhraseDescription;
                        if (contractData.TestTemplateEx.Authorization != null)
                        {
                            bool? b =
                                ValidateItem.ConvertToBoolean(Convert.ToChar(contractData.TestTemplateEx.Authorization));
                            row.Cells[4].Range.Text = b != null && (bool)b ? "כן" : "לא";
                        }


                        row.Cells[5].Range.Text = contractData.FinalPrice.ToString() + " ש" +
                                                  '\"' + "ח";
                    }

                    //word 2007 with 64 bit
                    SearchAndReplaceEverywhere(newDocument, "AAAA", DateTime.Now.ToShortDateString());
                    SearchAndReplaceEverywhere(newDocument, "BBBB", _currentClient.Name);
                    SearchAndReplaceEverywhere(newDocument, "CCCC", txtContactMan.Text);

                    //word 2003
                    //FindAndReplace(app.ActiveDocument.Content, "AAAA", DateTime.Now.ToShortDateString());
                    //FindAndReplace(app.ActiveDocument.Content, "BBBB", _selectedClient.Name);
                    //FindAndReplace(app.ActiveDocument.Content, "CCCC", txtContactMan.Text);
                    newDocument.SaveAs(savedFileName);
                    newDocument.Close();

                    try
                    {
                        WordToPdf.Convert(savedFileName, "docx");
                        btnOpenDoc.Enabled = true;
                    }
                    catch (Exception e)
                    {
                        CustomMessageBox.Show("המרה ל PDF נכשלה");
                        Logger.WriteLogFile(e);
                    }
                    MessageBox.Show("." + "יצירת מסמך סקר חוזה הושלמה", "Nautilus", MessageBoxButtons.OK,
                                   MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    //    app.Quit();
                    MessageBox.Show("." + "יצירת מסמך סקר חוזה נכשלה", "Nautilus", MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                    Logger.WriteLogFile(ex);
                }
                finally
                {
                    if (document != null) document.Close();
                    wordApp.Quit();
                }
            }
            else
            {
                MessageBox.Show("לא הוגדרה תיקיית יעד עבור הלקוח", "Nautilus");
            }
        }




        private string CreateSavedFileName(string templatefileName)
        {

            CombineClientPath();

            string savedFileName;
            if (!Directory.Exists(ClientPath))
            {
                CreateDirectory(ClientPath);

                //Create new file with new name
                savedFileName = CreateFileName(ClientPath, false);
            }

            else
            {
                //Create new file with new name
                savedFileName = CreateFileName(templatefileName, true);
            }
            return savedFileName;
        }


        private string CreateFileName(string templatefileName, bool hasLastDoc)
        {

            int max;
            if (hasLastDoc)
            {
                KeyValuePair<int, string> kvp = GetLastDoc("docx");
                max = kvp.Key + 1;
            }
            else
            {
                max = 1;
            }
            string date = "_" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;
            var validClientName = _currentClient.Name.MakeSafeFilename(' ').TrimEnd(' ');
            string fileName = Path.Combine(ClientPath, "_" + max.ToString() + "_" + validClientName + date + ".docx");

            return fileName;
        }
        private void CombineClientPath()
        {
            ClientPath = Path.Combine(clientDocumentsPath, txtPath.Text.MakeSafeFilename(' ').TrimEnd(' '));

        }
        private KeyValuePair<int, string> GetLastDoc(string suffix)
        {
            if (string.IsNullOrEmpty(ClientPath))
            {
                CombineClientPath();
            }
            //Get directory
            var dirInfo = new DirectoryInfo(ClientPath);
            if (!dirInfo.Exists)
            {
                CreateDirectory(ClientPath);
                dirInfo = new DirectoryInfo(ClientPath);
            }

            //Get word files from directory
            FileInfo[] files = dirInfo.GetFiles("*." + suffix);


            var temp = new Dictionary<int, string>(); //int is version + string is  name
            int max = 0;


            //If it is first time for this client
            if (files.Count() > 0)
            {
                foreach (FileInfo file in files)
                {
                    string item = file.Name;
                    string[] num = item.Split('_');
                    string n = num[1];
                    if (!temp.ContainsKey(int.Parse(n)))
                        temp.Add(int.Parse(n), item);
                }
                //Get max value from list Files versions
                max = temp.Keys.Max();
            }

            return temp.Where(x => x.Key == max).FirstOrDefault();
        }

        private void btnSendEmail_Click(object sender, EventArgs e)
        {
            try
            {
                var oApp = new Microsoft.Office.Interop.Outlook.Application();
                var oMailItem = (_MailItem)oApp.CreateItem(OlItemType.olMailItem);
                oMailItem.To = linkEmail.Text;

                string pdfName = GetLastDoc("pdf").Value;

                String sSource = pdfName;
                String sDisplayName = pdfName;
                //  int iPosition = oMailItem.Body.Length + 1;
                var iAttachType = (int)OlAttachmentType.olByValue;
                Attachment oAttach = oMailItem.Attachments.Add(ClientPath + sSource, iAttachType,
                    /*iPosition*/1, sDisplayName);
                oMailItem.Subject = string.Format("סקר חוזה {0} {1}", dropDownClients.Text, DateTime.Now);
                // body, bcc etc...
                oMailItem.Display(true);
            }
            catch (Exception exception)
            {
                MessageBox.Show("error");
                Logger.WriteLogFile(exception);
            }
        }

        private void btnOpenDoc_Click(object sender, EventArgs e)
        {
            string path = null;
            DateTime lastMod;
            try
            {


                KeyValuePair<int, string> kvp = GetLastDoc("docx");
                path = Path.Combine(ClientPath, kvp.Value);

                FileInfo fileInfo = new FileInfo(path);


                //תאריך שנוי אחרון של המסמך
                lastMod = fileInfo.LastWriteTime;


                Process p = Process.Start(path);
                p.WaitForExit();
            }
            catch (Exception exception)
            {

                One1.Controls.CustomMessageBox.Show("שגיאה בפתיחת המסמך", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.WriteLogFile(exception);
                //יציאה מהפונקציה אם לא נפתח המסמך
                return;
            }

            try
            {

                FileInfo newfileInfo = new FileInfo(path);

                //אם נעשו שינויים במסמך נייצר מסמך   
                //חדש pdf
                if (!DateTime.Equals(lastMod, newfileInfo.LastWriteTime))
                {
                    //Get pdf file
                    FileInfo pdfFile = new FileInfo(path.Replace(".docx", ".pdf"));

                    string otherPath = Path.Combine(pdfFile.Directory.ToString(), "(old)" + pdfFile.Name);
                    if (Directory.GetFiles(ClientPath, "*.pdf").Any(name => name == otherPath))
                        File.Copy(path, otherPath, true);
                    else
                        //Backup to old pdf
                        File.Move(pdfFile.FullName, otherPath);
                    //Convert new word to new pdf
                    WordToPdf.Convert(path, "docx");


                }

            }
            catch (Exception ex)
            {
                CustomMessageBox.Show("שגיאה ביצירת מסמך גבוי", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.WriteLogFile(ex);
            }
        }


        public void CreateDirectory(string path)
        {
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            catch (Exception e)
            {
                MessageBox.Show("יצירת ספריה נכשלה");
                Logger.WriteLogFile(e);
            }
            finally
            {
            }
        }

        #region WORD 2007 WITH 64 BIT

        private Document CopyToNewDocument(Document document)
        {
            document.StoryRanges[WdStoryType.wdMainTextStory].Copy();

            Document newDocument = document.Application.Documents.Add();
            newDocument.StoryRanges[WdStoryType.wdMainTextStory].Paste();
            return newDocument;
        }

        private void SearchAndReplaceEverywhere(
            Document document, string find, string replace)
        {
            foreach (Range storyRange in document.StoryRanges)
            {
                Range range = storyRange;
                while (range != null)
                {
                    SearchAndReplaceInStoryRange(range, find, replace);

                    if (range.ShapeRange.Count > 0)
                    {
                        foreach (Shape shape in range.ShapeRange)
                        {
                            if (shape.TextFrame.HasText != 0)
                            {
                                SearchAndReplaceInStoryRange(
                                    shape.TextFrame.TextRange, find, replace);
                            }
                        }
                    }
                    range = range.NextStoryRange;
                }
            }
        }

        private void SearchAndReplaceInStoryRange(
            Range range, string find, string replace)
        {
            range.Find.ClearFormatting();
            range.Find.Replacement.ClearFormatting();
            range.Find.Text = find;
            range.Find.Replacement.Text = replace;
            range.Find.Wrap = WdFindWrap.wdFindContinue;
            range.Find.Execute(Replace: WdReplace.wdReplaceAll);
        }

        #endregion

        #region Word 2003

        private void FindAndReplace(Range range, object findText, object replaceText)
        {
            object item = WdGoToItem.wdGoToPage;
            object whichItem = WdGoToDirection.wdGoToFirst;
            object replaceAll = WdReplace.wdReplaceAll;
            object forward = true;
            object matchAllWord = true;
            object missing = Missing.Value;

            range.Document.GoTo(ref item, ref whichItem, ref missing, ref missing);
            range.Find.Execute(ref findText, ref missing, ref matchAllWord,
                               ref missing, ref missing, ref missing, ref forward,
                               ref missing, ref missing, ref replaceText, ref replaceAll,
                               ref missing, ref missing, ref missing, ref missing);
        }

        #endregion

        #endregion


        private void textBoxRemarks_Leave(object sender, EventArgs e)
        {
            try
            {
                lastContract.U_REMARKS = textBoxRemarks.Text;
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
            }
        }

        private void txtDiscount_Leave(object sender, EventArgs e)
        {
            try
            {
                decimal discount;
                var b = decimal.TryParse(txtDiscount.Text, out discount);

                _currentClient.Discount = discount;
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
            }
        }









    }
}
//        private void AddAutoCompleteTextBox()
//        {
//            txtSearchTest.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
//
//            RadListDataItemCollection autoCompleteItems = txtSearchTest.AutoCompleteItems;
//
//            //Clear previous complete items 
//            autoCompleteItems.Clear();
//
//            string tag = radPageLabs.SelectedPage.Tag.ToString();
//
//            if (tag == "ALL")
//            {
//                autoCompleteItems.AddRange(_testTemplateList.Select(x => x.Name).ToList());
//            }
//            else
//            {
//                //Get Test template names  for current tab
//                IEnumerable<string> testTemplateNames =
//                    _testTemplateList.Where(x => x.In_labs != null && x.In_labs.Contains(tag)).Select(t => t.Name);
//
//                autoCompleteItems.AddRange(testTemplateNames);
//            }
//        }
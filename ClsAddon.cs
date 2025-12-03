using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Xml;
using System.Xml.Linq;


namespace SBAAddon
{
    public class ClsAddon
    {
        public static SAPbouiCOM.Application SBO_Application = null;
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Form form, oForm, targetForm, sbaForm, activeSOForm, soForm = null;
        public static SAPbouiCOM.Matrix oMatrix, Matrix, mtx, sbaMatrix = null;
        private static bool isSOAlreadyOpened = false;
        private static bool isSQAlreadyOpened = false;
        public static string StatusMessage = "";
        public static int count = 0;
        string _oldValue = "";
        string _oldCol = "";
        int _oldRow = -1;

        public ClsAddon()
        {
            SetApplication();
            clsUtilities clsUtility = new clsUtilities();
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
        }

        #region SetApplication
        private void SetApplication()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();
                string ConnectionString = (System.Environment.GetCommandLineArgs().Length == 0) ? System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0)) : System.Environment.GetCommandLineArgs().GetValue(1).ToString();
                SboGuiApi.Connect(ConnectionString);
                SBO_Application = SboGuiApi.GetApplication(-1);
                oCompany = SBO_Application.Company.GetDICompany();
                SBO_Application.StatusBar.SetSystemMessage("Addon is initialized with " + oCompany.CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception Ex)
            {
                SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == BoAppEventTypes.aet_ShutDown || EventType == BoAppEventTypes.aet_CompanyChanged || EventType == BoAppEventTypes.aet_LanguageChanged || EventType == BoAppEventTypes.aet_ServerTerminition)
                System.Windows.Forms.Application.Exit();
        }

        #region LoadFromXML Defination
        public static void LoadFromXML(string FileName)
        {
            try
            {
                XmlDocument oXmlDoc = new XmlDocument();
                string sPath;
                SAPbouiCOM.Form sboForm = SBO_Application.Forms.GetFormByTypeAndCount(169, 1);
                sboForm.Freeze(true);
                try
                {
                    SAPbobsCOM.Recordset PTS_Rec;
                    sPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\XMLFiles\\" + FileName.ToString();
                    oXmlDoc.Load(sPath);
                    SBO_Application.LoadBatchActions(oXmlDoc.InnerXml);
                    PTS_Rec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    PTS_Rec.DoQuery("SELECT \"Color\" FROM oadm");
                    switch (PTS_Rec.RecordCount)
                    {
                        case 0: SBO_Application.ActivateMenuItem("47649"); break;
                    }
                    Marshal.ReleaseComObject(PTS_Rec);
                    GC.Collect();
                }
                finally
                {
                    sboForm.Freeze(false);
                    sboForm.Update();
                }
            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }
        }
        public static string LoadFromXML_FormLoad(SAPbouiCOM.Application app, out bool BubbleEvent, string FileName, string DocEntry)
        {
            string FrmUID;
            XmlNode oXNode;
            XmlAttribute oAttr;
            XmlDocument oXmlDoc = new XmlDocument();
            BubbleEvent = true;
            try
            {
                Random r = new Random();
                string sFilePath = System.Windows.Forms.Application.StartupPath.ToString() + "\\XMLFiles\\" + FileName;
                oXmlDoc.Load(sFilePath);
                string xmlString = oXmlDoc.InnerText.ToString();
                oXNode = oXmlDoc.GetElementsByTagName("form").Item(0);
                oAttr = (XmlAttribute)oXNode.Attributes.GetNamedItem("uid");
                oAttr.Value = oAttr.Value + r.Next(111, 999).ToString();
                FrmUID = oAttr.Value;
                app.LoadBatchActions(oXmlDoc.InnerXml);
                return FrmUID;
            }
            catch (Exception ex)
            {

                BubbleEvent = false;
                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return "";
            }
            finally
            {
                oXmlDoc = null;
            }
        }
        #endregion
        #endregion

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region Add  Buttons  
                if (pVal.FormTypeEx == "1250000100" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Item item = oForm.Items.Add("BtnPrint", BoFormItemTypes.it_BUTTON);
                    item.Left = oForm.Items.Item("1250000002").Left + 80;
                    item.Width = 100;
                    item.Top = oForm.Items.Item("1250000002").Top;
                    item.Height = oForm.Items.Item("1250000002").Height;
                    ((SAPbouiCOM.Button)item.Specific).Caption = "Copy To SO";

                    SAPbouiCOM.Item qItem = oForm.Items.Add("BtnQuot", BoFormItemTypes.it_BUTTON);
                    qItem.Left = item.Left + item.Width + 10;
                    qItem.Width = 100;
                    qItem.Top = item.Top;
                    qItem.Height = item.Height;
                    ((SAPbouiCOM.Button)qItem.Specific).Caption = "Copy To SQ";
                }

                if (pVal.FormTypeEx == "139" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Item item = oForm.Items.Item("2");
                    SAPbouiCOM.Item btnItem = oForm.Items.Add("btnSBA", BoFormItemTypes.it_BUTTON);
                    btnItem.Left = item.Left + item.Width + 10;
                    btnItem.Top = item.Top;
                    btnItem.Width = 90;
                    btnItem.Height = item.Height;
                    ((SAPbouiCOM.Button)btnItem.Specific).Caption = "Select SBA";
                }
                #endregion

                #region Sales blanket Agreement Item Category Combo Select
                if (pVal.FormTypeEx == "1250000100" && pVal.ItemUID == "U_ItemCategory" && pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE && !pVal.BeforeAction && pVal.ActionSuccess)
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.ComboBox cmbCat = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_ItemCategory").Specific;
                    string selectedCategory = cmbCat.Value;
                    string discountText = ((SAPbouiCOM.EditText)oForm.Items.Item("U_Discount").Specific).Value;
                    double Discount = string.IsNullOrWhiteSpace(discountText) ? 0 : Convert.ToDouble(discountText);

                    if (!string.IsNullOrEmpty(selectedCategory))
                    {
                        string cardCode = ((SAPbouiCOM.EditText)oForm.Items.Item("1250000006").Specific).Value;
                        string query1 = $@"SELECT T0.""ItemCode"", T0.""ItemName"", T1.""Price"", T2.""ItmsGrpNam""  FROM OITM T0  INNER JOIN OCRD C ON C.""CardCode"" = '{cardCode}'  INNER JOIN ITM1 T1 ON T1.""ItemCode"" = T0.""ItemCode"" AND T1.""PriceList"" = C.""ListNum"" INNER JOIN OITB T2 ON T0.""ItmsGrpCod"" = T2.""ItmsGrpCod""  WHERE T0.""U_ItemCategory"" = '{selectedCategory}'  AND T0.""frozenFor"" = 'N' ";
                        Recordset oRS = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRS.DoQuery(query1);
                        oForm.PaneLevel = 2;
                        SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)oForm.Items.Item("1250000045").Specific;

                        mtx.Clear();
                        try
                        {
                            oForm.Freeze(true);
                            int currentRow = 1;
                            while (!oRS.EoF)
                            {
                                if (mtx.RowCount < currentRow)
                                    mtx.AddRow();
                                string itemCode = oRS.Fields.Item("ItemCode").Value.ToString();
                                double actprice = Convert.ToDouble(oRS.Fields.Item("Price").Value);
                                SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                mtx.Columns.Item("1250000001").Cells.Item(currentRow).Specific.Value = itemCode;
                                mtx.Columns.Item("1250000007").Cells.Item(currentRow).Specific.Value = "1";
                                double finalPrice = 0.00;
                                if (Discount > 0.00)
                                    finalPrice = actprice - Discount;

                                if (finalPrice < 0.00)
                                    finalPrice = 0;

                                mtx.Columns.Item("1250000009").Cells.Item(currentRow).Specific.Value = finalPrice.ToString("F2");
                                //   mtx.Columns.Item("U_DiscInRs").Cells.Item(currentRow).Specific.Value = Discount.ToString();
                                mtx.Columns.Item("U_UTL_DISCPER").Cells.Item(currentRow).Specific.Value = Discount.ToString();

                                currentRow++;
                                oRS.MoveNext();
                            }
                            mtx.AutoResizeColumns();
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }
                }
                #endregion

                #region Sales blanket Agreement PriceList Combo Select
                if (pVal.FormTypeEx == "1250000100" && pVal.ItemUID == "1250000045" && pVal.ColUID == "U_PriceList" && pVal.EventType == BoEventTypes.et_COMBO_SELECT && !pVal.BeforeAction && pVal.ActionSuccess)
                {
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    string discountText = ((SAPbouiCOM.EditText)oForm.Items.Item("U_Discount").Specific).Value;
                    double discount = string.IsNullOrWhiteSpace(discountText) ? 0 : Convert.ToDouble(discountText);

                    SAPbouiCOM.Matrix sbaMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1250000045").Specific;
                    int row = pVal.Row;

                    SAPbouiCOM.ComboBox cmbPriceList = (SAPbouiCOM.ComboBox)sbaMatrix.Columns.Item("U_PriceList").Cells.Item(row).Specific;
                    string selectedPriceList = cmbPriceList.Selected == null ? "" : cmbPriceList.Selected.Value;
                    string itemCode = ((SAPbouiCOM.EditText)sbaMatrix.Columns.Item("1250000001").Cells.Item(row).Specific).Value;

                    if (!string.IsNullOrEmpty(selectedPriceList) && !string.IsNullOrEmpty(itemCode))
                    {
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $"SELECT T1.\"Price\" FROM OITM T0 INNER JOIN ITM1 T1 ON T0.\"ItemCode\" = T1.\"ItemCode\" WHERE T1.\"PriceList\" = '{selectedPriceList}' AND T0.\"ItemCode\" = '{itemCode}'";
                        oRS.DoQuery(query);

                        if (!oRS.EoF)
                        {
                            double basePrice = Convert.ToDouble(oRS.Fields.Item("Price").Value);
                            double finalPrice = basePrice;

                            if (discount > 0.00)
                                finalPrice = basePrice - discount;

                            if (finalPrice < 0.00)
                                finalPrice = 0;

                            ((SAPbouiCOM.EditText)sbaMatrix.Columns.Item("1250000009").Cells.Item(row).Specific).Value = finalPrice.ToString("F2");
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                    }
                }
                #endregion

                #region Sales blanket Agreement CopyTO SO
                if (pVal.FormTypeEx == "1250000100" && pVal.ItemUID == "BtnPrint" && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK || pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && !pVal.BeforeAction)
                {
                    try
                    {
                        oForm = SBO_Application.Forms.Item(pVal.FormUID);

                        var cmbStatus = (SAPbouiCOM.ComboBox)oForm.Items.Item("1250000036").Specific;
                        string docStatus = cmbStatus.Selected?.Value ?? "";
                        if (docStatus != "A")
                        {
                            SBO_Application.StatusBar.SetText("Action not allowed. Document must be Approved.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }

                        DateTime endDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("1250000016").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                        if (endDate < DateTime.Today)
                        {
                            SBO_Application.StatusBar.SetText("End Date has already passed. Cannot set this date for the sales order.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        var cmbBranch = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_BPLId").Specific;
                        var cmbItemCategory = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_ItemCategory").Specific;
                        string agreementNo = ((SAPbouiCOM.EditText)oForm.Items.Item("1250000004").Specific).Value;
                        string bpCode = ((SAPbouiCOM.EditText)oForm.Items.Item("1250000006").Specific).Value;

                        string branchCode = cmbBranch.Selected?.Value ?? "";
                        string itemCategory = cmbItemCategory.Selected?.Value ?? "";

                        if (string.IsNullOrEmpty(branchCode) || string.IsNullOrEmpty(itemCategory))
                        {
                            SBO_Application.StatusBar.SetText("Branch Code and Item Category must be selected before proceeding.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        SAPbouiCOM.Matrix sbaMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1250000045").Specific;
                        DateTime formStartDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("1250000014").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        if (isSOAlreadyOpened) return;
                        isSOAlreadyOpened = true;

                        for (int i = 1; i <= sbaMatrix.RowCount; i++)
                        {
                            var chk = (SAPbouiCOM.CheckBox)sbaMatrix.Columns.Item("U_Select").Cells.Item(i).Specific;
                            if (!chk.Checked) continue;
                            string itemCode = ((SAPbouiCOM.EditText)sbaMatrix.Columns.Item("1250000001").Cells.Item(i).Specific).Value;
                            var rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string qt = $@"SELECT T0.""Number"" AS SBA_No, (T0.""U_INKGS"" - T0.""U_ConInKGS"") AS OpenQty, T0.""StartDate"",  T0.""EndDate"", T0.""TermDate"",T0.""Status"" FROM ""OOAT"" T0 INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE  T0.""BpCode"" = '{bpCode}' AND T0.""U_ItemCategory"" = '{itemCategory}' AND T0.""U_BPLId"" = '{branchCode}' AND T1.""ItemCode"" = '{itemCode}'  AND (T0.""U_INKGS"" - T0.""U_ConInKGS"") > 0  AND T0.""Status""  = 'A' AND (T0.""TermDate"" IS NULL OR T0.""TermDate"" > CURRENT_DATE)  AND T0.""EndDate"" >= CURRENT_DATE AND CURRENT_DATE BETWEEN T0.""StartDate"" AND COALESCE(T0.""TermDate"", T0.""EndDate"") ORDER BY  T0.""StartDate"" ASC";
                            rs.DoQuery(qt);

                            if (rs.RecordCount > 0)
                            {
                                rs.MoveFirst();
                                string openSBA = rs.Fields.Item("SBA_No").Value.ToString();
                                double openQty = Convert.ToDouble(rs.Fields.Item("OpenQty").Value);
                                DateTime sbaStartDate = Convert.ToDateTime(rs.Fields.Item("StartDate").Value);
                                DateTime sbaEndDate = Convert.ToDateTime(rs.Fields.Item("EndDate").Value);

                                if (sbaEndDate < DateTime.Today)
                                {
                                    SBO_Application.StatusBar.SetText($"Cannot set values to Sales Order. SBA {openSBA} ended on {sbaEndDate:yyyy-MM-dd}.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    isSOAlreadyOpened = false;
                                    return;
                                }

                                if (sbaStartDate <= formStartDate && openQty > 0.0001 && agreementNo != openSBA)
                                {
                                    SBO_Application.StatusBar.SetText($"Cannot set values to Sales Order. SBA {openSBA} is already open from {sbaStartDate:yyyy-MM-dd}.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    isSOAlreadyOpened = false;
                                    return;
                                }
                            }
                        }

                        SBO_Application.ActivateMenuItem("2050");
                        SAPbouiCOM.Form soForm = SBO_Application.Forms.ActiveForm;
                        if (soForm.TypeEx != "139")
                        {
                            SBO_Application.StatusBar.SetText("Could not open Sales Order form.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            isSOAlreadyOpened = false;
                            return;
                        }

                        ((SAPbouiCOM.EditText)soForm.Items.Item("4").Specific).Value = bpCode;
                        SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        if (!string.IsNullOrEmpty(branchCode))
                        {
                            var soBranch = (SAPbouiCOM.ComboBox)soForm.Items.Item("2001").Specific;
                            soBranch.Select(branchCode, BoSearchKey.psk_ByValue);
                        }

                        SAPbouiCOM.Matrix soMatrix = (SAPbouiCOM.Matrix)soForm.Items.Item("38").Specific;
                        try
                        {
                            soForm.Freeze(true);
                            soMatrix.Clear();
                            var dirs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string di = $@" SELECT *   FROM ""OOAT"" T0  INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo""  WHERE T0.""Number"" = '{agreementNo}'  AND T0.""BpCode"" = '{bpCode}'   AND T0.""U_ItemCategory"" = '{itemCategory}'  AND T0.""U_BPLId"" = '{branchCode}'  AND T1.""U_Select"" = 'Y'";
                            dirs.DoQuery(di);

                            int targetRow = 1;
                            while (!dirs.EoF)
                            {
                                if (targetRow > soMatrix.RowCount)
                                {
                                    soMatrix.AddRow();
                                }
                                string itemCode = Convert.ToString(dirs.Fields.Item("ItemCode").Value);
                                double lineQty = Convert.ToDouble(dirs.Fields.Item("PlanQty").Value);
                                double actPrice = Convert.ToDouble(dirs.Fields.Item("UnitPrice").Value);
                                SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                ((SAPbouiCOM.EditText)soMatrix.Columns.Item("1").Cells.Item(targetRow).Specific).Value = itemCode;
                                ((SAPbouiCOM.EditText)soMatrix.Columns.Item("1980002193").Cells.Item(targetRow).Specific).Value = agreementNo.Trim();
                                ((SAPbouiCOM.EditText)soMatrix.Columns.Item("11").Cells.Item(targetRow).Specific).Value = lineQty.ToString("0.00");
                                ((SAPbouiCOM.EditText)soMatrix.Columns.Item("14").Cells.Item(targetRow).Specific).Value = actPrice.ToString("0.00");
                                targetRow++;
                                dirs.MoveNext();
                            }

                          ((SAPbouiCOM.EditText)soForm.Items.Item("16").Specific).Value = "Origin: Blanket Agreement: " + agreementNo;
                            isSOAlreadyOpened = false;
                        }
                        finally
                        {
                            soForm.Freeze(false);
                        }

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        isSOAlreadyOpened = false;
                        BubbleEvent = false;
                    }
                }
                #endregion

                #region Sales blanket Agreement CopyTO SQ
                if (pVal.FormTypeEx == "1250000100" && pVal.ItemUID == "BtnQuot" && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK || pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) && !pVal.BeforeAction)
                {
                    try
                    {
                        oForm = SBO_Application.Forms.Item(pVal.FormUID);

                        var cmbStatus = (SAPbouiCOM.ComboBox)oForm.Items.Item("1250000036").Specific;
                        string docStatus = cmbStatus.Selected?.Value ?? "";
                        if (docStatus != "A")
                        {
                            SBO_Application.StatusBar.SetText("Action not allowed. Document must be Approved.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }

                        DateTime endDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("1250000016").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                        if (endDate < DateTime.Today)
                        {
                            SBO_Application.StatusBar.SetText("End Date has already passed. Cannot set this date for the sales quotation.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        var cmbBranch = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_BPLId").Specific;
                        var cmbItemCategory = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_ItemCategory").Specific;
                        string agreementNo = ((SAPbouiCOM.EditText)oForm.Items.Item("1250000004").Specific).Value;
                        string bpCode = ((SAPbouiCOM.EditText)oForm.Items.Item("1250000006").Specific).Value;

                        string branchCode = cmbBranch.Selected?.Value ?? "";
                        string itemCategory = cmbItemCategory.Selected?.Value ?? "";

                        if (string.IsNullOrEmpty(branchCode) || string.IsNullOrEmpty(itemCategory))
                        {
                            SBO_Application.StatusBar.SetText("Branch Code and Item Category must be selected before proceeding.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        SAPbouiCOM.Matrix sbaMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1250000045").Specific;
                        DateTime formStartDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("1250000014").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        if (isSQAlreadyOpened) return;
                        isSQAlreadyOpened = true;

                        for (int i = 1; i <= sbaMatrix.RowCount; i++)
                        {
                            var chk = (SAPbouiCOM.CheckBox)sbaMatrix.Columns.Item("U_Select").Cells.Item(i).Specific;
                            if (!chk.Checked) continue;
                            string itemCode = ((SAPbouiCOM.EditText)sbaMatrix.Columns.Item("1250000001").Cells.Item(i).Specific).Value;
                            var rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string qt = $@"SELECT T0.""Number"" AS SBA_No, (T0.""U_INKGS"" - T0.""U_ConInKGS"") AS OpenQty, T0.""StartDate"",  T0.""EndDate"", T0.""TermDate"",T0.""Status"" FROM ""OOAT"" T0 INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE  T0.""BpCode"" = '{bpCode}' AND T0.""U_ItemCategory"" = '{itemCategory}' AND T0.""U_BPLId"" = '{branchCode}' AND T1.""ItemCode"" = '{itemCode}'  AND (T0.""U_INKGS"" - T0.""U_ConInKGS"") > 0  AND T0.""Status""  = 'A' AND (T0.""TermDate"" IS NULL OR T0.""TermDate"" > CURRENT_DATE)  AND T0.""EndDate"" >= CURRENT_DATE AND CURRENT_DATE BETWEEN T0.""StartDate"" AND COALESCE(T0.""TermDate"", T0.""EndDate"") ORDER BY  T0.""StartDate"" ASC";
                            rs.DoQuery(qt);

                            if (rs.RecordCount > 0)
                            {
                                rs.MoveFirst();
                                string openSBA = rs.Fields.Item("SBA_No").Value.ToString();
                                double openQty = Convert.ToDouble(rs.Fields.Item("OpenQty").Value);
                                DateTime sbaStartDate = Convert.ToDateTime(rs.Fields.Item("StartDate").Value);
                                DateTime sbaEndDate = Convert.ToDateTime(rs.Fields.Item("EndDate").Value);

                                if (sbaEndDate < DateTime.Today)
                                {
                                    SBO_Application.StatusBar.SetText($"Cannot set values to Sales Quotation. SBA {openSBA} ended on {sbaEndDate:yyyy-MM-dd}.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    isSQAlreadyOpened = false;
                                    return;
                                }

                                if (sbaStartDate <= formStartDate && openQty > 0.0001 && agreementNo != openSBA)
                                {
                                    SBO_Application.StatusBar.SetText($"Cannot set values to Sales Quotation. SBA {openSBA} is already open from {sbaStartDate:yyyy-MM-dd}.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    isSQAlreadyOpened = false;
                                    return;
                                }
                            }
                        }

                        SBO_Application.ActivateMenuItem("2048");
                        SAPbouiCOM.Form sqForm = SBO_Application.Forms.ActiveForm;
                        if (sqForm.TypeEx != "149")
                        {
                            SBO_Application.StatusBar.SetText("Could not open Sales Quotation form.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            isSQAlreadyOpened = false;
                            return;
                        }

                        ((SAPbouiCOM.EditText)sqForm.Items.Item("4").Specific).Value = bpCode;
                        SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        if (!string.IsNullOrEmpty(branchCode))
                        {
                            var sqBranch = (SAPbouiCOM.ComboBox)sqForm.Items.Item("2001").Specific;
                            sqBranch.Select(branchCode, BoSearchKey.psk_ByValue);
                        }

                        SAPbouiCOM.Matrix sqMatrix = (SAPbouiCOM.Matrix)sqForm.Items.Item("38").Specific;
                        try
                        {
                            sqForm.Freeze(true);
                            sqMatrix.Clear();
                            var dirs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string di = $@" SELECT *   FROM ""OOAT"" T0  INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE T0.""Number"" = '{agreementNo}'  AND T0.""BpCode"" = '{bpCode}'   AND T0.""U_ItemCategory"" = '{itemCategory}'  AND T0.""U_BPLId"" = '{branchCode}'  AND T1.""U_Select"" = 'Y'";
                            dirs.DoQuery(di);

                            int targetRow = 1;
                            while (!dirs.EoF)
                            {
                                if (targetRow > sqMatrix.RowCount)
                                {
                                    sqMatrix.AddRow();
                                }
                                string itemCode = Convert.ToString(dirs.Fields.Item("ItemCode").Value);
                                double lineQty = Convert.ToDouble(dirs.Fields.Item("PlanQty").Value);
                                double actPrice = Convert.ToDouble(dirs.Fields.Item("UnitPrice").Value);
                                SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                ((SAPbouiCOM.EditText)sqMatrix.Columns.Item("1").Cells.Item(targetRow).Specific).Value = itemCode;
                                ((SAPbouiCOM.EditText)sqMatrix.Columns.Item("1980002193").Cells.Item(targetRow).Specific).Value = agreementNo.Trim();
                                ((SAPbouiCOM.EditText)sqMatrix.Columns.Item("11").Cells.Item(targetRow).Specific).Value = lineQty.ToString("0.00");
                                ((SAPbouiCOM.EditText)sqMatrix.Columns.Item("14").Cells.Item(targetRow).Specific).Value = actPrice.ToString("0.00");
                                targetRow++;
                                dirs.MoveNext();
                            }

                            ((SAPbouiCOM.EditText)sqForm.Items.Item("16").Specific).Value = "Origin: Blanket Agreement: " + agreementNo;
                            isSQAlreadyOpened = false;
                        }
                        finally
                        {
                            sqForm.Freeze(false);
                        }

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        isSQAlreadyOpened = false;
                        BubbleEvent = false;
                    }
                }
                #endregion

                #region Sales Order btnSBA
                if (pVal.FormTypeEx == "139" && pVal.ItemUID == "btnSBA" && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form soForm = SBO_Application.Forms.Item(pVal.FormUID);
                    string cardCode = soForm.DataSources.DBDataSources.Item("ORDR").GetValue("CardCode", 0).Trim();
                    if (string.IsNullOrWhiteSpace(cardCode))
                    {
                        SBO_Application.StatusBar.SetText("CardCode not found on Sales Order.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    LoadFromXML_FormLoad(SBO_Application, out BubbleEvent, "Sales Blanket Agreement Details.xml", "");
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.ActiveForm;
                    SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MAT").Specific;

                    SAPbouiCOM.DBDataSource oDBS_Details = oForm.DataSources.DBDataSources.Item("@SOSBA");

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    // string query = $@"SELECT DISTINCT T0.""AbsID"" AS ""DocEntry"", T0.""Number"" AS ""DocNum"", T0.""BpCode"", T0.""StartDate"", T0.""EndDate"",T0.""U_ItemCategory"", T0.""Type"", T0.""Status"", T0.""U_BPLId"" FROM OOAT T0  INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo""   WHERE T0.""BpCode"" = '{cardCode.Replace("'", "''")}'  AND T0.""Status"" = 'A'  AND T1.""U_Select"" = 'Y'  AND CURRENT_DATE BETWEEN T0.""StartDate"" AND T0.""EndDate""ORDER BY T0.""Number""";
                    string query = $@"
                                       SELECT DISTINCT 
                                           T0.""AbsID"" AS ""DocEntry"", 
                                           T0.""Number"" AS ""DocNum"", 
                                           T0.""BpCode"", 
                                           T0.""StartDate"", 
                                           T0.""EndDate"",
                                           T0.""U_ItemCategory"", 
                                           T0.""Type"", 
                                           T0.""Status"", 
                                           T0.""U_BPLId""

                                       FROM OOAT T0
                                       INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo""
                                       WHERE T0.""BpCode"" = '{cardCode.Replace("'", "''")}'
                                         AND T0.""Status"" = 'A'
                                         AND T1.""U_Select"" = 'Y'
                                         AND (T0.""U_INKGS"" - T0.""U_ConInKGS"") > 0
                                         AND (T0.""TermDate"" IS NULL OR T0.""TermDate"" >= CURRENT_DATE)
                                         AND T0.""EndDate"" >= CURRENT_DATE
                                         AND CURRENT_DATE BETWEEN T0.""StartDate"" AND COALESCE(T0.""TermDate"", T0.""EndDate"")
                                       ORDER BY T0.""Number""";


                    rs.DoQuery(query);

                    oForm.Freeze(true);
                    matrix.Clear();
                    int targetRow = 1;
                    while (!rs.EoF)
                    {
                        if (targetRow > matrix.RowCount)
                            matrix.AddRow();

                        string AbsID = rs.Fields.Item("DocEntry").Value.ToString();
                        string Number = rs.Fields.Item("DocNum").Value.ToString();
                        string CustomerCode = rs.Fields.Item("BpCode").Value.ToString();
                        string ItemCategory = rs.Fields.Item("U_ItemCategory").Value.ToString();
                        string AgrType = rs.Fields.Item("Type").Value.ToString();
                        string Status = rs.Fields.Item("Status").Value.ToString();
                        string Branch = rs.Fields.Item("U_BPLId").Value.ToString();

                        DateTime StartDate = DateTime.Parse(rs.Fields.Item("StartDate").Value.ToString());
                        DateTime EndDate = DateTime.Parse(rs.Fields.Item("EndDate").Value.ToString());
                        SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        matrix.Columns.Item("#").Cells.Item(targetRow).Specific.Value = targetRow.ToString();
                        matrix.Columns.Item("Col_10").Cells.Item(targetRow).Specific.Value = AbsID;
                        matrix.Columns.Item("Col_11").Cells.Item(targetRow).Specific.Value = Number;
                        matrix.Columns.Item("Col_1").Cells.Item(targetRow).Specific.Value = CustomerCode;
                        matrix.Columns.Item("Col_2").Cells.Item(targetRow).Specific.Value = ItemCategory;
                        matrix.Columns.Item("Col_3").Cells.Item(targetRow).Specific.Value = StartDate.ToString("yyyyMMdd");
                        matrix.Columns.Item("Col_4").Cells.Item(targetRow).Specific.Value = EndDate.ToString("yyyyMMdd");
                        matrix.Columns.Item("Col_5").Cells.Item(targetRow).Specific.Value = AgrType;
                        matrix.Columns.Item("Col_6").Cells.Item(targetRow).Specific.Value = Status;
                        matrix.Columns.Item("Col_7").Cells.Item(targetRow).Specific.Value = Branch;

                        targetRow++;
                        rs.MoveNext();
                    }
                    matrix.AutoResizeColumns();
                    oForm.Freeze(false);
                }
                #endregion

                #region Set data of custom form to sales order 
                if (pVal.FormTypeEx == "SOSBA" && pVal.ItemUID == "Submit" && pVal.EventType == BoEventTypes.et_CLICK && !pVal.BeforeAction && pVal.ActionSuccess)
                {
                    SAPbouiCOM.Form sbaForm = SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Matrix sbaMatrix = (SAPbouiCOM.Matrix)sbaForm.Items.Item("MAT").Specific;

                    List<string> selectedAbsIds = new List<string>();
                    string commonBranch = null;
                    bool branchMismatch = false;

                    for (int i = 1; i <= sbaMatrix.RowCount; i++)
                    {
                        SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)sbaMatrix.Columns.Item("Col_0").Cells.Item(i).Specific;
                        if (!chk.Checked) continue;

                        string absId = ((SAPbouiCOM.EditText)sbaMatrix.Columns.Item("Col_10").Cells.Item(i).Specific).Value;
                        string sqlBranch = $@"SELECT ""U_BPLId"" FROM OOAT WHERE ""AbsID"" = '{absId}'";
                        SAPbobsCOM.Recordset rsBranch = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsBranch.DoQuery(sqlBranch);
                        string bplId = rsBranch.Fields.Item("U_BPLId").Value.ToString();

                        if (commonBranch == null)
                            commonBranch = bplId;
                        else if (commonBranch != bplId)
                        {
                            branchMismatch = true;
                            break;
                        }

                        selectedAbsIds.Add(absId);
                    }

                    if (branchMismatch)
                    {
                        SBO_Application.StatusBar.SetText("Selected SBAs belong to different branches. Please select SBAs from the same branch.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    SAPbouiCOM.Form soForm = SBO_Application.Forms.Cast<SAPbouiCOM.Form>().FirstOrDefault(f => f.TypeEx == "139");
                    if (soForm == null)
                    {
                        SBO_Application.StatusBar.SetText("Sales Order form not open.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    if (!string.IsNullOrEmpty(commonBranch))
                    {
                        SAPbouiCOM.ComboBox soBranch = (SAPbouiCOM.ComboBox)soForm.Items.Item("2001").Specific;
                        soBranch.Select(commonBranch, BoSearchKey.psk_ByValue);
                    }

                    SAPbouiCOM.Matrix soMatrix = (SAPbouiCOM.Matrix)soForm.Items.Item("38").Specific;
                    soForm.Freeze(true);
                    soMatrix.Clear();

                    string remarkText = "Origin: Blanket Agreement: ";
                    int row = 1;

                    foreach (string absId in selectedAbsIds)
                    {
                        string sql = $@"SELECT H.""Number"", H.""U_BPLId"", L.""ItemCode"", L.""PlanQty"", L.""UnitPrice""   FROM OOAT H INNER JOIN OAT1 L ON H.""AbsID"" = L.""AgrNo""   WHERE H.""AbsID"" = '{absId}' AND L.""U_Select"" = 'Y'";
                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rs.DoQuery(sql);

                        while (!rs.EoF)
                        {
                            if (row > soMatrix.RowCount)
                                soMatrix.AddRow();
                            SBO_Application.StatusBar.SetText("Please wait... Data loading in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                            ((SAPbouiCOM.EditText)soMatrix.Columns.Item("1").Cells.Item(row).Specific).Value = rs.Fields.Item("ItemCode").Value.ToString();
                            ((SAPbouiCOM.EditText)soMatrix.Columns.Item("1980002193").Cells.Item(row).Specific).Value = rs.Fields.Item("Number").Value.ToString();
                            ((SAPbouiCOM.EditText)soMatrix.Columns.Item("11").Cells.Item(row).Specific).Value = Convert.ToDouble(rs.Fields.Item("PlanQty").Value).ToString("0.00");
                            ((SAPbouiCOM.EditText)soMatrix.Columns.Item("14").Cells.Item(row).Specific).Value = Convert.ToDouble(rs.Fields.Item("UnitPrice").Value).ToString("0.00");
                            //((SAPbouiCOM.EditText)soMatrix.Columns.Item("234000525").Cells.Item(row).Specific).Value = Convert.ToDouble(rs.Fields.Item("UnitPrice").Value).ToString("0.00");
                            //((SAPbouiCOM.EditText)soMatrix.Columns.Item("234000526").Cells.Item(row).Specific).Value = Convert.ToDouble(rs.Fields.Item("UnitPrice").Value).ToString("0.00");

                            remarkText += rs.Fields.Item("Number").Value.ToString() + "; ";
                            row++;
                            rs.MoveNext();
                        }
                    }

                    ((SAPbouiCOM.EditText)soForm.Items.Item("16").Specific).Value = remarkText.TrimEnd(';', ' ');
                    soForm.Freeze(false);
                    sbaForm.Close();
                }
                #endregion

                #region Validate Qty before Add (Only for Add Mode)  sales order
                if (pVal.FormTypeEx == "139" && (pVal.EventType == BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_ITEM_PRESSED) && pVal.ItemUID == "1" && pVal.BeforeAction == true)
                {
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                            string bpcode = oForm.Items.Item("4").Specific.value;
                            SAPbouiCOM.ComboBox cmbBranch = (SAPbouiCOM.ComboBox)oForm.Items.Item("2001").Specific;
                            string branchCode = cmbBranch.Selected?.Value ?? "";

                            Dictionary<string, double> baConsumptionMap = new Dictionary<string, double>();
                            Dictionary<string, double> baMaxAllowedMap = new Dictionary<string, double>();

                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value.Trim();
                                string agrNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1980002193").Cells.Item(i).Specific).Value.Trim();
                                double qty = 0;
                                double.TryParse(((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value.Trim(), out qty);

                                if (!string.IsNullOrEmpty(itemCode) && !string.IsNullOrEmpty(agrNo) && qty > 0.00)
                                {
                                    SAPbobsCOM.Recordset rsItem = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rsItem.DoQuery($@"SELECT COALESCE(""IWeight1"",0) AS ""Weight"" FROM OITM WHERE ""ItemCode""='{itemCode}'");
                                    double weight = !rsItem.EoF ? Convert.ToDouble(rsItem.Fields.Item("Weight").Value) : 0;

                                    if (weight <= 0)
                                    {
                                        BubbleEvent = false;
                                        SBO_Application.StatusBar.SetText($"Error: Item '{itemCode}' does not have a valid weight in Inventory.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        return;
                                    }

                                    double lineWeight = qty * weight;
                                    if (!baConsumptionMap.ContainsKey(agrNo))
                                    {
                                        SAPbobsCOM.Recordset rsSBA = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rsSBA.DoQuery($@"SELECT COALESCE(""U_ConInKGS"",0) AS ""Consumed"", COALESCE(""U_INKGS"",0) AS ""MaxAllowed"" FROM ""OOAT""   WHERE ""Number"" = {agrNo} AND ""BpCode"" = '{bpcode}' AND ""U_BPLId"" = '{branchCode}'");

                                        double alreadyConsumed = 0;
                                        double maxAllowed = 0;

                                        if (!rsSBA.EoF)
                                        {
                                            alreadyConsumed = Convert.ToDouble(rsSBA.Fields.Item("Consumed").Value);
                                            maxAllowed = Convert.ToDouble(rsSBA.Fields.Item("MaxAllowed").Value);
                                        }

                                        baConsumptionMap[agrNo] = alreadyConsumed + lineWeight;
                                        baMaxAllowedMap[agrNo] = maxAllowed;
                                    }
                                    else
                                    {
                                        baConsumptionMap[agrNo] += lineWeight;
                                    }

                                    if (baConsumptionMap[agrNo] > baMaxAllowedMap[agrNo])
                                    {
                                        BubbleEvent = false;
                                        SBO_Application.StatusBar.SetText($"SO exceeds Blanket Agreement {agrNo}. Allowed: {baMaxAllowedMap[agrNo]}, Attempted: {baConsumptionMap[agrNo]}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            BubbleEvent = false;
                            SBO_Application.StatusBar.SetText($"Unexpected error while validating SO: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                #endregion

                #region Sales blanket Agreement Validate Blanket Agreement Discounts
                if (pVal.FormTypeEx == "1250000100" && pVal.BeforeAction == true && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1250000001")
                {
                    try
                    {
                        SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);

                        string headerDiscStr = oForm.DataSources.DBDataSources.Item("OOAT").GetValue("U_Discount", 0).Trim();
                        double headerDisc = 0;
                        double.TryParse(headerDiscStr, out headerDisc);

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1250000045").Specific;
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            string rowDiscStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_UTL_DISCPER").Cells.Item(i).Specific).Value.Trim();
                            // string rowDiscStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_DiscInRs").Cells.Item(i).Specific).Value.Trim();
                            double rowDisc = 0;
                            double.TryParse(rowDiscStr, out rowDisc);

                            if (rowDisc > headerDisc)
                            {
                                BubbleEvent = false;
                                SBO_Application.StatusBar.SetText($"Cannot add Blanket Agreement: Row discount ({rowDisc}%) exceeds header discount ({headerDisc}%).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                return;
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        BubbleEvent = false;
                        SBO_Application.StatusBar.SetText($"Error while validating Blanket Agreement: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }

                if (pVal.FormTypeEx == "1250000100" && pVal.BeforeAction == true && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1250000001")
                {
                    try
                    {
                        SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);
                        int absId = 0;
                        string absIdStr = oForm.DataSources.DBDataSources.Item("OOAT").GetValue("AbsID", 0).Trim();
                        int.TryParse(absIdStr, out absId);

                        if (absId > 0)
                        {
                            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                            string sqlHeader = $@"SELECT COALESCE(""U_Discount"",0) AS ""HeaderDiscount""   FROM ""OOAT"" WHERE ""AbsID"" = {absId}";
                            rs.DoQuery(sqlHeader);
                            double headerDisc = !rs.EoF ? Convert.ToDouble(rs.Fields.Item("HeaderDiscount").Value) : 0;

                            string sqlRows = $@"SELECT COALESCE(""U_UTL_DISCPER"",0) AS ""RowDisc"" FROM ""OAT1"" WHERE ""AgrNo"" = {absId}";
                            //  string sqlRows = $@"SELECT COALESCE(""U_DiscInRs"",0) AS ""RowDisc"" FROM ""OAT1"" WHERE ""AgrNo"" = {absId}";
                            rs.DoQuery(sqlRows);

                            while (!rs.EoF)
                            {
                                double rowDisc = Convert.ToDouble(rs.Fields.Item("RowDisc").Value);
                                if (rowDisc > headerDisc)
                                {
                                    BubbleEvent = false;
                                    SBO_Application.StatusBar.SetText($"Cannot update Blanket Agreement: Row discount ({rowDisc}%) exceeds header discount ({headerDisc}%).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    return;
                                }
                                rs.MoveNext();
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        BubbleEvent = false;
                        SBO_Application.StatusBar.SetText($"Error while validating Blanket Agreement: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region Prevent Editing Price or Quantity in Existing Sales Orders

                if (pVal.FormTypeEx == "139" && pVal.ItemUID == "38")
                {
                    if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
                    {
                        oForm = SBO_Application.Forms.Item(FormUID);
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                        _oldCol = pVal.ColUID;
                        _oldRow = pVal.Row;

                        if (_oldCol == "11" || _oldCol == "14")
                        {
                            _oldValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(_oldCol).Cells.Item(_oldRow).Specific).Value.Trim();
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                    {
                        oForm = SBO_Application.Forms.Item(FormUID);
                        if (oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            string col = pVal.ColUID;
                            if ((col == "11" || col == "14") && pVal.Row == _oldRow)
                            {
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                                string newValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(col).Cells.Item(pVal.Row).Specific).Value.Trim();

                                if (newValue != _oldValue)
                                {
                                    SBO_Application.StatusBar.SetText("Updating Price or Quantity is not allowed on existing Sales Orders.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                isSOAlreadyOpened = false;
            }
        }

        private void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region sales order after add 
                if (BusinessObjectInfo.FormTypeEx == "139" && BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess)
                {
                    int docEntry = 0;
                    if (!string.IsNullOrEmpty(BusinessObjectInfo.ObjectKey))
                    {
                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                        docEntry = int.Parse(xmlDoc.SelectSingleNode("DocumentParams/DocEntry").InnerText);
                    }

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string sql = $@"SELECT T1.""AgrNo"" AS ""AgreementNo"", SUM(COALESCE(T1.""Quantity"",0) * COALESCE(OITM.""IWeight1"",0)) AS ""TotalWeight""  FROM RDR1 T1     INNER JOIN OITM ON T1.""ItemCode"" = OITM.""ItemCode""  WHERE T1.""DocEntry"" = {docEntry} AND T1.""AgrNo"" IS NOT NULL  GROUP BY T1.""AgrNo"";  ";
                    rs.DoQuery(sql);

                    while (!rs.EoF)
                    {
                        string sbaNumber = rs.Fields.Item("AgreementNo").Value.ToString();
                        double totalWeight = Convert.ToDouble(rs.Fields.Item("TotalWeight").Value);

                        SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string checkSql = $@"SELECT COALESCE(""U_ConInKGS"",0) AS ""Consumed"", COALESCE(""U_INKGS"",0) AS ""MaxAllowed""  FROM ""OOAT"" WHERE ""AbsID"" = {sbaNumber}";
                        rsCheck.DoQuery(checkSql);

                        double alreadyConsumed = 0;
                        double maxAllowed = 0;
                        if (!rsCheck.EoF)
                        {
                            alreadyConsumed = Convert.ToDouble(rsCheck.Fields.Item("Consumed").Value);
                            maxAllowed = Convert.ToDouble(rsCheck.Fields.Item("MaxAllowed").Value);
                        }

                        double newConsumed = alreadyConsumed;
                        newConsumed += totalWeight;

                        if (newConsumed < 0)
                            newConsumed = 0;

                        if (newConsumed != alreadyConsumed)
                        {
                            SAPbobsCOM.Recordset rsUpd = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string upd = $@"UPDATE ""OOAT"" SET ""U_ConInKGS"" = {newConsumed} WHERE ""AbsID"" = {sbaNumber};";
                            rsUpd.DoQuery(upd);
                        }
                        rs.MoveNext();
                    }

                }
                #endregion

                #region  sales order cancel update  
                if (BusinessObjectInfo.FormTypeEx == "139" && BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE && BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                {
                    int docEntry = 0;
                    try
                    {
                        if (!string.IsNullOrEmpty(BusinessObjectInfo.ObjectKey))
                        {
                            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                            xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                            docEntry = int.Parse(xmlDoc.SelectSingleNode("DocumentParams/DocEntry").InnerText);
                        }

                        if (docEntry == 0)
                            return;

                        SAPbobsCOM.Documents oOrder = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        if (!oOrder.GetByKey(docEntry) || oOrder.Cancelled != SAPbobsCOM.BoYesNoEnum.tYES)
                            return;

                        string processedFlag = oOrder.UserFields.Fields.Item("U_Process").Value.ToString();
                        if (processedFlag == "Y")
                        {
                            SBO_Application.StatusBar.SetText($"Sales Order {docEntry} cancellation already processed.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            return;
                        }

                        SAPbobsCOM.Recordset rsLines = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string query = $@"SELECT ""AgrNo"", ""Quantity"", ""ItemCode"" FROM RDR1 WHERE ""DocEntry"" = {docEntry}";
                        rsLines.DoQuery(query);

                        while (!rsLines.EoF)
                        {
                            int agrNo = Convert.ToInt32(rsLines.Fields.Item("AgrNo").Value);
                            if (agrNo == 0)
                            {
                                rsLines.MoveNext();
                                continue;
                            }

                            double qty = Convert.ToDouble(rsLines.Fields.Item("Quantity").Value);
                            string itemCode = rsLines.Fields.Item("ItemCode").Value.ToString();

                            SAPbobsCOM.Items oItem =
                                (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                            double weight = 0;

                            if (oItem.GetByKey(itemCode))
                            {
                                object weightObj = oItem.UserFields.Fields.Item("IWeight1").Value;
                                weight = weightObj != null && weightObj.ToString() != "" ? Convert.ToDouble(weightObj) : 0;
                            }

                            double reduction = qty * weight;
                            string updateSql = $@"UPDATE OOAT  SET ""U_ConInKGS"" = COALESCE(""U_ConInKGS"", 0) - {reduction} WHERE ""AbsID"" = {agrNo}";
                            rsUpdate.DoQuery(updateSql);

                            rsLines.MoveNext();
                        }

                        oOrder.UserFields.Fields.Item("U_Process").Value = "Y";
                        int res = oOrder.Update();
                        if (res != 0)
                        {
                            oCompany.GetLastError(out int errCode, out string errMsg);
                            throw new Exception(errMsg);
                        }

                        SBO_Application.StatusBar.SetText($"Sales Order {docEntry} cancelled — OOAT updated once.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText($"Error on Sales Order cancel: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


    }
}
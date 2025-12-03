using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SBAAddon
{
    public class clsDBStructure
    {
        public clsDBStructure()
        {
        }

        public static void addField(string TblName, string Title, string Desc, string Type, int Length, string Structure, string LinkTable)
        {
            bool flag = false;
            Recordset businessObject = (Recordset)((dynamic)ClsAddon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            string[] title = new string[] { "select \"FieldID\", \"EditSize\" from CUFD where \"AliasID\" ='", Title, "' and \"TableID\" ='", TblName, "'" };
            businessObject.DoQuery(string.Concat(title));
            int value = (int)((dynamic)businessObject.Fields.Item("FieldID").Value);
            int num = (int)((dynamic)businessObject.Fields.Item("EditSize").Value);
            if (businessObject.RecordCount == 0)
            {
                flag = true;
            }
            Marshal.ReleaseComObject(businessObject);
            GC.Collect();
            UserFieldsMD tblName = (UserFieldsMD)((dynamic)ClsAddon.oCompany.GetBusinessObject(BoObjectTypes.oUserFields));
            if (flag)
            {
                tblName.TableName = TblName;
                tblName.Name = Title;
                tblName.Description = Desc;
                tblName.LinkedTable = LinkTable;
                if (Type == "A" && Structure == "A")
                {
                    tblName.Type = BoFieldTypes.db_Alpha;
                    tblName.Size = Length;
                    if (tblName.TableName == "INV1" && tblName.Name == "BType")
                    {
                        tblName.ValidValues.Value = "HIRE";
                        tblName.ValidValues.Description = "HIRE";
                        tblName.ValidValues.Add();
                        tblName.ValidValues.Value = "OT";
                        tblName.ValidValues.Description = "OT";
                        tblName.ValidValues.Add();
                    }



                    if (tblName.TableName == "ORDR" && tblName.Name == "Process")
                    {
                        tblName.ValidValues.Value = "Y";
                        tblName.ValidValues.Description = "YES";
                        tblName.ValidValues.Add();
                        tblName.ValidValues.Value = "N";
                        tblName.ValidValues.Description = "NO";
                        tblName.ValidValues.Add();
                        tblName.DefaultValue = "N";
                    }


                }
                if (Type == "A" && Structure == "?")
                {
                    tblName.Type = BoFieldTypes.db_Alpha;
                }
                if (Type == "A" && Structure == "#")
                {
                    tblName.Type = BoFieldTypes.db_Alpha;
                }
                if (Type == "A" && Structure == "M")
                {
                    tblName.Type = BoFieldTypes.db_Alpha;
                }
                if (Type == "N" && Structure == "N")
                {
                    tblName.Type = BoFieldTypes.db_Numeric;
                    tblName.EditSize = Length;
                }
                if (Type == "D" && Structure == "D")
                {
                    tblName.Type = BoFieldTypes.db_Date;
                }
                if (Type == "D" && Structure == "T")
                {
                    tblName.Type = BoFieldTypes.db_Date;
                }
                if (Type == "B" && Structure == "R")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "B" && Structure == "S")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "B" && Structure == "P")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "B" && Structure == "Q")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "B" && Structure == "%")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "B" && Structure == "M")
                {
                    tblName.Type = BoFieldTypes.db_Float;
                }
                if (Type == "M" && Structure == "B")
                {
                    tblName.Type = BoFieldTypes.db_Memo;
                }
                if (Type == "R" && Structure == "I")
                {
                    tblName.Type = BoFieldTypes.db_Alpha;
                }
                string structure = Structure;
                string str = structure;
                if (structure != null)
                {
                    switch (str)
                    {
                        case "A":
                            {
                                tblName.SubType = BoFldSubTypes.st_None;
                                break;
                            }
                        case "?":
                            {
                                tblName.SubType = BoFldSubTypes.st_Address;
                                break;
                            }
                        case "#":
                            {
                                tblName.SubType = BoFldSubTypes.st_Phone;
                                break;
                            }
                        case "N":
                            {
                                tblName.SubType = BoFldSubTypes.st_None;
                                break;
                            }
                        case "D":
                            {
                                tblName.SubType = BoFldSubTypes.st_None;
                                break;
                            }
                        case "T":
                            {
                                tblName.SubType = BoFldSubTypes.st_Time;
                                break;
                            }
                        case "R":
                            {
                                tblName.SubType = BoFldSubTypes.st_Rate;
                                break;
                            }
                        case "S":
                            {
                                tblName.SubType = BoFldSubTypes.st_Sum;
                                break;
                            }
                        case "P":
                            {
                                tblName.SubType = BoFldSubTypes.st_Price;
                                break;
                            }
                        case "Q":
                            {
                                tblName.SubType = BoFldSubTypes.st_Quantity;
                                break;
                            }
                        case "%":
                            {
                                tblName.SubType = BoFldSubTypes.st_Percentage;
                                break;
                            }
                        case "M":
                            {
                                tblName.SubType = BoFldSubTypes.st_Measurement;
                                break;
                            }
                        case "B":
                            {
                                tblName.SubType = BoFldSubTypes.st_Link;
                                break;
                            }
                        case "I":
                            {
                                tblName.SubType = BoFldSubTypes.st_Image;
                                break;
                            }
                    }
                }
                if (tblName.Add() != 0)
                {
                    ClsAddon.SBO_Application.StatusBar.SetText(string.Concat(TblName, Title, ClsAddon.oCompany.GetLastErrorDescription()), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    ClsAddon.SBO_Application.StatusBar.SetText(string.Concat(Title, " Created Sucessfully in ", TblName), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }
            Marshal.ReleaseComObject(tblName);
            GC.Collect();
        }

        public static void addTable(string tblName, string tblDsc, int tblType)
        {
            UserTablesMD businessObject = (UserTablesMD)((dynamic)ClsAddon.oCompany.GetBusinessObject(BoObjectTypes.oUserTables));
            if (!businessObject.GetByKey(tblName))
            {
                businessObject.TableName = tblName;
                businessObject.TableDescription = tblDsc;
                switch (tblType)
                {
                    case 0:
                        {
                            businessObject.TableType = BoUTBTableType.bott_NoObject;
                            break;
                        }
                    case 1:
                        {
                            businessObject.TableType = BoUTBTableType.bott_MasterData;
                            break;
                        }
                    case 2:
                        {
                            businessObject.TableType = BoUTBTableType.bott_MasterDataLines;
                            break;
                        }
                    case 3:
                        {
                            businessObject.TableType = BoUTBTableType.bott_Document;
                            break;
                        }
                    case 4:
                        {
                            businessObject.TableType = BoUTBTableType.bott_DocumentLines;
                            break;
                        }
                    case 5:
                        {
                            businessObject.TableType = BoUTBTableType.bott_NoObjectAutoIncrement;
                            break;
                        }
                }
                if (businessObject.Add() != 0)
                {
                    ClsAddon.SBO_Application.StatusBar.SetText(ClsAddon.oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    ClsAddon.SBO_Application.StatusBar.SetText(string.Concat(tblName, " Created Sucessfully."), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }
            Marshal.ReleaseComObject(businessObject);
            businessObject = null;
            GC.Collect();
        }

        public static void RegisterUDO(string UDO_Code, string UDO_Name, int Tbl_Type, string Tbl_Name, string Menu_Caption, int Position, string Def_Fields, string Child_Tbl,  string Def_Form)
        {
            UserObjectsMD businessObject = (UserObjectsMD)((dynamic)ClsAddon.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD));
            if (!businessObject.GetByKey(UDO_Code))
            {
                businessObject.Code = UDO_Code;
                businessObject.Name = UDO_Name;
                businessObject.TableName = Tbl_Name;
                switch (Tbl_Type)
                {
                    case 1:
                        {
                            businessObject.ObjectType = BoUDOObjType.boud_MasterData;
                            goto case 2;
                        }
                    case 2:
                        {
                            businessObject.CanFind = BoYesNoEnum.tYES;
                            businessObject.CanDelete = BoYesNoEnum.tYES;
                            businessObject.CanCancel = BoYesNoEnum.tYES;
                            businessObject.CanLog = BoYesNoEnum.tYES;
                            businessObject.CanClose = BoYesNoEnum.tYES;
                            businessObject.ManageSeries = BoYesNoEnum.tYES;
                            if (Def_Fields != "" && Def_Form == "Y")
                            {
                                businessObject.CanCreateDefaultForm = BoYesNoEnum.tYES;
                                businessObject.CanFind = BoYesNoEnum.tNO;
                                businessObject.EnableEnhancedForm = BoYesNoEnum.tNO;
                                string[] strArrays = Def_Fields.Split(new char[] { ',' });
                                for (int i = 0; i < (int)strArrays.Length; i++)
                                {
                                    string str = strArrays[i];
                                    string[] strArrays1 = str.Split(new char[] { '|' });
                                    if ((int)strArrays1.Length > 0)
                                    {
                                        if (strArrays1[0].ToString() == "Code" || strArrays1[0].ToString() == "Name" || strArrays1[0].ToString() == "DocNum" || strArrays1[0].ToString() == "DocEntry")
                                        {
                                            businessObject.FormColumns.FormColumnAlias = strArrays1[0].ToString();
                                        }
                                        else
                                        {
                                            businessObject.FormColumns.FormColumnAlias = string.Concat("U_", strArrays1[0].ToString());
                                        }
                                    }
                                    if ((int)strArrays1.Length > 1)
                                    {
                                        businessObject.FormColumns.FormColumnDescription = strArrays1[0].ToString();
                                    }
                                    businessObject.FormColumns.Add();
                                }
                            }
                            if (businessObject.CanFind == BoYesNoEnum.tYES)
                            {
                                if (businessObject.ObjectType != BoUDOObjType.boud_Document)
                                {
                                    businessObject.FindColumns.ColumnAlias = "Code";
                                    businessObject.FindColumns.ColumnDescription = "Code";
                                    businessObject.FindColumns.Add();
                                }
                                else
                                {
                                    businessObject.FindColumns.ColumnAlias = "DocNum";
                                    businessObject.FindColumns.ColumnDescription = "DocNum";
                                    businessObject.FindColumns.Add();
                                    businessObject.FindColumns.ColumnAlias = "CreateDate";
                                    businessObject.FindColumns.ColumnDescription = "CreateDate";
                                    businessObject.FindColumns.Add();
                                    businessObject.FindColumns.ColumnAlias = "UpdateDate";
                                    businessObject.FindColumns.ColumnDescription = "UpdateDate";
                                    businessObject.FindColumns.Add();
                                }
                            }
                            if (Child_Tbl != "")
                            {
                                string[] strArrays2 = Child_Tbl.Split(new char[] { '/' });
                                for (int j = 0; j < (int)strArrays2.Length; j++)
                                {
                                    businessObject.ChildTables.TableName = strArrays2[j];
                                    businessObject.ChildTables.Add();
                                }
                            }
                            if (businessObject.Add() != 0)
                            {
                                ClsAddon.SBO_Application.StatusBar.SetText(ClsAddon.oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                ClsAddon.SBO_Application.StatusBar.SetText(string.Concat(Tbl_Name, " Created Sucessfully."), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            }
                            Marshal.ReleaseComObject(businessObject);
                            businessObject = null;
                            GC.Collect();
                            break;
                        }
                    case 3:
                        {
                            businessObject.ObjectType = BoUDOObjType.boud_Document;
                            goto case 2;
                        }
                    default:
                        {
                            goto case 2;
                        }
                }
            }
        }

        public static void RegisterUDO1(string UDO_Code, string UDO_Name, int Tbl_Type, string Tbl_Name, string Menu_Caption, int Position, string Def_Fields, string Child_Tbl, string Child_Tbl2, string Def_Form)
        {
            UserObjectsMD businessObject = (UserObjectsMD)((dynamic)ClsAddon.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD));
            if (!businessObject.GetByKey(UDO_Code))
            {
                businessObject.Code = UDO_Code;
                businessObject.Name = UDO_Name;
                businessObject.TableName = Tbl_Name;

                switch (Tbl_Type)
                {
                    case 1:
                        businessObject.ObjectType = BoUDOObjType.boud_MasterData;
                        goto case 2;

                    case 2:
                        businessObject.CanFind = BoYesNoEnum.tYES;
                        businessObject.CanDelete = BoYesNoEnum.tYES;
                        businessObject.CanCancel = BoYesNoEnum.tYES;
                        businessObject.CanLog = BoYesNoEnum.tYES;
                        businessObject.CanClose = BoYesNoEnum.tYES;
                        businessObject.ManageSeries = BoYesNoEnum.tYES;

                        if (Def_Fields != "" && Def_Form == "Y")
                        {
                            businessObject.CanCreateDefaultForm = BoYesNoEnum.tYES;
                            businessObject.CanFind = BoYesNoEnum.tNO;
                            businessObject.EnableEnhancedForm = BoYesNoEnum.tNO;

                            string[] strArrays = Def_Fields.Split(',');
                            for (int i = 0; i < strArrays.Length; i++)
                            {
                                string[] strArrays1 = strArrays[i].Split('|');
                                if (strArrays1.Length > 0)
                                {
                                    if (strArrays1[0] == "Code" || strArrays1[0] == "Name" || strArrays1[0] == "DocNum" || strArrays1[0] == "DocEntry")
                                        businessObject.FormColumns.FormColumnAlias = strArrays1[0];
                                    else
                                        businessObject.FormColumns.FormColumnAlias = "U_" + strArrays1[0];
                                }
                                if (strArrays1.Length > 1)
                                {
                                    businessObject.FormColumns.FormColumnDescription = strArrays1[0];
                                }
                                businessObject.FormColumns.Add();
                            }
                        }

                        if (businessObject.CanFind == BoYesNoEnum.tYES)
                        {
                            if (businessObject.ObjectType != BoUDOObjType.boud_Document)
                            {
                                businessObject.FindColumns.ColumnAlias = "Code";
                                businessObject.FindColumns.ColumnDescription = "Code";
                                businessObject.FindColumns.Add();
                            }
                            else
                            {
                                businessObject.FindColumns.ColumnAlias = "DocNum";
                                businessObject.FindColumns.ColumnDescription = "DocNum";
                                businessObject.FindColumns.Add();

                                businessObject.FindColumns.ColumnAlias = "CreateDate";
                                businessObject.FindColumns.ColumnDescription = "CreateDate";
                                businessObject.FindColumns.Add();

                                businessObject.FindColumns.ColumnAlias = "UpdateDate";
                                businessObject.FindColumns.ColumnDescription = "UpdateDate";
                                businessObject.FindColumns.Add();
                            }
                        }

                        // ✅ Add first child table
                        if (!string.IsNullOrEmpty(Child_Tbl))
                        {
                            string[] strArrays2 = Child_Tbl.Split('/');
                            for (int j = 0; j < strArrays2.Length; j++)
                            {
                                businessObject.ChildTables.TableName = strArrays2[j];
                                businessObject.ChildTables.Add();
                            }
                        }

                        // ✅ Add second child table
                        if (!string.IsNullOrEmpty(Child_Tbl2))
                        {
                            string[] strArrays3 = Child_Tbl2.Split('/');
                            for (int j = 0; j < strArrays3.Length; j++)
                            {
                                businessObject.ChildTables.TableName = strArrays3[j];
                                businessObject.ChildTables.Add();
                            }
                        }

                        if (businessObject.Add() != 0)
                        {
                            ClsAddon.SBO_Application.StatusBar.SetText(ClsAddon.oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            ClsAddon.SBO_Application.StatusBar.SetText(Tbl_Name + " Created Successfully.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }

                        Marshal.ReleaseComObject(businessObject);
                        businessObject = null;
                        GC.Collect();
                        break;

                    case 3:
                        businessObject.ObjectType = BoUDOObjType.boud_Document;
                        goto case 2;

                    default:
                        goto case 2;
                }
            }
        }


    }

}

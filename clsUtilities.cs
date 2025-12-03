using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;


namespace SBAAddon
{
    public class clsUtilities
    {
        public clsUtilities()
        {
            clsDBStructure.addField("ORDR", "Process", "Process Done", "A", 10, "A", "");

            clsDBStructure.addField("OOAT", "Qty", "Quantity", "B", 0, "Q", "");
            clsDBStructure.addField("OOAT", "ConInKGS", "Consumed InKGS", "B", 0, "Q", "");

            // "DiscInRs", "Discount in Rs" 
            // "DiscInRs", "Discount in Rs" 

            clsDBStructure.addTable("SOSBA", "Sales Blankent Details", 1);
            clsDBStructure.addField("@SOSBA", "CustomerCode", "Customer Code", "A", 80, "", "");
            clsDBStructure.addField("@SOSBA", "ItemCategory", "Item Category", "A", 80, "", "");
            clsDBStructure.addField("@SOSBA", "StartDate", "Start Date", "D", 10, "D", "");
            clsDBStructure.addField("@SOSBA", "EndDate", "End Date", "D", 10, "D", "");
            clsDBStructure.addField("@SOSBA", "AgrType", "Agreement Type", "A", 80, "", "");
            clsDBStructure.addField("@SOSBA", "Status", "Status", "A", 80, "A", "");
            clsDBStructure.addField("@SOSBA", "Branch", "Branch", "A", 80, "A", "");
            clsDBStructure.addField("@SOSBA", "select", "select", "A", 10, "A", "");
            clsDBStructure.addField("@SOSBA", "ItemCode", "Item Code", "A", 80, "A", "");
            clsDBStructure.addField("@SOSBA", "ItemDesc", "Item Description", "A", 100, "A", "");
            clsDBStructure.addField("@SOSBA", "AbsID", "Item Code", "A", 80, "A", "");
            clsDBStructure.addField("@SOSBA", "Number", "Item Description", "A", 100, "A", "");

            clsDBStructure.RegisterUDO("SOSBA", "Sales Blankent Details", 1, "SOSBA", "Sales Blankent Details", 0, "", "", "");

        }
    }

}
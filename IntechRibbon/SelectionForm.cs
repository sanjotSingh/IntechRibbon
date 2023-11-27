using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;



namespace IntechRibbon
{
    public partial class SelectionForm : System.Windows.Forms.Form
    {



        public SelectionForm()
        {
            InitializeComponent();
            this.CenterToParent();
            //add each schedule here

            //checkedListBox.Items.Add("excel");
        }

        private void checkedListBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, true);
            }
        }

        private void uncheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, false);
            }
        }

        private void toggle_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, !checkedListBox.GetItemChecked(i));
            }
        }

        private void bomExport_Click(object sender, EventArgs e)
        {
            this.Close(); //just closing the form
        }
    }
}

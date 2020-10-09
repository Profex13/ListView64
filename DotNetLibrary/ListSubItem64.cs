using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;

namespace DotNetLibrary
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IListSubItem64
    {
        [DispId(0x10000001)]
        bool Bold { get; set; }

        //[DispId(0x10000002)]
        //Object ForeColor { get; set; }
    
        //[DispId(0x10000003)]
        //int Index { get; set; }
    
        [DispId(0x10000004)]
        string Key { get; set; }
          
        //[DispId(0x10000005)]
        //dynamic ReportIcon { get; set; }

        [DispId(0x10000006)]
        dynamic Tag { get; set; }

        [DispId(0x10000007)]
        string Text { get; set; }
            
        //[DispId(0x10000008)]
        //string ToolTipText { get; set; }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [DefaultProperty("Text")]
    public class ListSubItem64 : System.Windows.Forms.ListViewItem.ListViewSubItem, IListSubItem64
    {
        private System.Windows.Forms.ListViewItem.ListViewSubItem _listViewSubItem;

        public ListSubItem64() : base()
        {            
            _listViewSubItem = new System.Windows.Forms.ListViewItem.ListViewSubItem();
        }

        public ListSubItem64(System.Windows.Forms.ListViewItem.ListViewSubItem ListViewSubItem)        
        {
            _listViewSubItem = ListViewSubItem;
        }

        public bool Bold
        {
            get { return (_listViewSubItem.Font.Bold); }
            set { _listViewSubItem.Font = new Font(_listViewSubItem.Font, value ? FontStyle.Bold : FontStyle.Regular); }
        }        

        public string Key 
        { 
            get { return (_listViewSubItem.Name); }
            set { _listViewSubItem.Name = value; }
        }

        public new dynamic Tag
        { 
            get { return (_listViewSubItem.Tag); }
            set { _listViewSubItem.Tag = value; }
        }

        public new string Text
        {
            get { return (_listViewSubItem.Text); }
            set { _listViewSubItem.Text = value; }
        }
    }
}

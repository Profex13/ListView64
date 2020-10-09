using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace DotNetLibrary
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IColumnHeader64Collection
    {
        [DispId(0x10000001)]
        ColumnHeader64 Add(int Index = -1, string Key = "", string Text = "", double Width = 0.0);

        [DispId(0x10000002)]
        void Clear();

        [DispId(0x10000003)]
        int Count { get; }
        
        [DispId(0x10000004)]
        bool IsReadOnly { get; }

        [DispId(0x10000005)]
        ColumnHeader64 this[int Index] { get; }       //Item
    }

    //[ProgId("DotNetLibrary.ColumnHeaders")]
    [ComVisible(true)]
    //[ClassInterface(ClassInterfaceType.AutoDual)]
    [ClassInterface(ClassInterfaceType.None)]
    //[ComSourceInterfaces(typeof(IListView64Events))]
    public class ColumnHeader64Collection : ICollection<ColumnHeader64>, IEnumerable, IColumnHeader64Collection
    {
        public ColumnHeader64Collection(System.Windows.Forms.ListView.ColumnHeaderCollection Columns) 
        { 
            this._columns = Columns;
        }
        private System.Windows.Forms.ListView.ColumnHeaderCollection _columns;

        public IEnumerator<ColumnHeader64> GetEnumerator()
        {
            foreach (System.Windows.Forms.ColumnHeader columnHeader in this._columns)
            {
                ColumnHeader64 column64 = new ColumnHeader64(columnHeader);
                if (column64 != null)
                    yield return column64;
            }
        }

        System.Collections.IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        } 

        public void Add(ColumnHeader64 Item)
        {
            _columns.Add(Item.Text);
        }

        public ColumnHeader64 Add(int Index = -1, string Key = "", string Text = "", double Width = 0.0)
        {            
            int width = (int)(Width / 0.75);
            //return new ColumnHeader64(_columns.Add(Text,Width));
            if (Index >= 0)
            {
                _columns.Insert(Index, Key, Text, width);
                return new ColumnHeader64(_columns[Index]);
            }
            else            
                if (Width == 0.0)                
                    if (Key == "")                    
                        return new ColumnHeader64(_columns.Add(Text));
                    else
                        return new ColumnHeader64(_columns.Add(Key, Text));                
                else
                    return new ColumnHeader64(_columns.Add(Key,Text,width));            
        }

        public void Clear()
        {
            _columns.Clear();
        }

        public int Count
        {
            get { return _columns.Count; }
        }

        public bool IsReadOnly
        {
            get { return _columns.IsReadOnly; }
        }

        public bool Contains(ColumnHeader64 item)
        {
            return _columns.Contains((System.Windows.Forms.ColumnHeader)item);
        }

        public void CopyTo(ColumnHeader64[] array, int arrayindex)
        {            
        }        

        public ColumnHeader64 this[int Index]       //Item
        {
            get { return new ColumnHeader64(_columns[Index]); }
        }

        //public ColumnHeader64 this[string Key]      //Item
        //{
        //    get { return (ColumnHeader64)_columns[Key]; }
        //}

        public bool Remove(ColumnHeader64 Column)
        {            
            try 
            {                
                _columns.Remove((System.Windows.Forms.ColumnHeader)Column);
            }
            catch(Exception ex)
            {                                
                MessageBox.Show(ex.ToString());
                return false;
            }
            return true;
        }
    }
}
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


//Don't import the entire namespace, this will cause name conflicts.
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlWin = Microsoft.Office.Interop.Excel.Window;

namespace ExcelExtensions
{

    public partial class ExcelAppCollection
    {

        #region Methods

        public static xlApp InnerFromProcess(Process p)
        {
            return InnerFromHandle(ChildHandleFromMainHandle(p.MainWindowHandle.ToInt32()));
        }

        public static Int32 ChildHandleFromMainHandle(Int32 mainHandle)
        {
            Int32 handle = 0;
            EnumChildWindows(mainHandle, EnumChildFunc, ref handle);
            return handle;
        }

        public static xlApp InnerFromHandle(Int32 handle)
        {
            xlWin win = null;
            Int32 hr = AccessibleObjectFromWindow(handle, DW_OBJECTID, rrid.ToByteArray(), ref win);
            return win.Application;
        }

        public static Int32 GetWindowZ(IntPtr handle)
        {
            var z = 0;
            for (IntPtr h = handle; h != IntPtr.Zero; h = GetWindow(h, GW_HWNDPREV))
                z++;
            return z;
        }

        public static Boolean EnumChildFunc(Int32 hwndChild, ref Int32 lParam)
        {
            var buf = new StringBuilder(128);
            GetClassName(hwndChild, buf, 128);
            if (buf.ToString() == ComClassName)
            {
                lParam = hwndChild;
                return false;
            }
            return true;
        }

        #endregion

        #region Extern Methods

        [DllImport("Oleacc.dll")]
        public static extern Int32 AccessibleObjectFromWindow(
            Int32 hwnd, UInt32 dwObjectID, Byte[] riid, ref xlWin ptr);

        [DllImport("User32.dll")]
        public static extern Boolean EnumChildWindows(
            Int32 hWndParent, EnumChildCallback lpEnumFunc, ref Int32 lParam);

        [DllImport("User32.dll")]
        public static extern Int32 GetClassName(
            Int32 hWnd, StringBuilder lpClassName, Int32 nMaxCount);

        [DllImport("User32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, UInt32 uCmd);

        #endregion

        #region Constants & delegates

        public const String MarshalName = "Excel.Application";

        public const String ProcessName = "EXCEL";

        public const String ComClassName = "EXCEL7";

        public const UInt32 DW_OBJECTID = 0xFFFFFFF0;

        public const UInt32 GW_HWNDPREV = 3;
        //3 = GW_HWNDPREV
        //The retrieved handle identifies the window above the specified window in the Z order.
        //If the specified window is a topmost window, the handle identifies a topmost window.
        //If the specified window is a top-level window, the handle identifies a top-level window.
        //If the specified window is a child window, the handle identifies a sibling window.

        public static Guid rrid = new Guid("{00020400-0000-0000-C000-000000000046}");

        public delegate Boolean EnumChildCallback(Int32 hwnd, ref Int32 lParam);
        #endregion
    }
}

namespace ExcelInstanceLoader
{
    internal class Utils
    {
        public static Dictionary<int, bool> GetExcelProcessIds()
        {
            // Get a list of all running processes
            Process[] processes = Process.GetProcessesByName("EXCEL");

            Dictionary<int, bool> processIds = new Dictionary<int, bool>();
            foreach (Process process in processes)
            {
                processIds[process.Id] = process.MainWindowTitle.Length == 0;
            }

            return processIds;
        }

        public static List<string> GetWorkbookNamesByProcessId(int processId)
        {
            List<string> workbookNames = new List<string>();

            xlApp excelApp = null;

            try
            {
                Process process = Process.GetProcessById(processId);
                excelApp = ExcelExtensions.ExcelAppCollection.InnerFromProcess(process);
            }
            catch (Exception ex)
            {
                // Handle exception if Excel is not running or no workbooks are open
                Console.WriteLine(ex.Message);
                return workbookNames;
            }

            foreach (Excel.Workbook workbook in excelApp.Workbooks)
            {
                workbookNames.Add(workbook.Name);
            }

            Marshal.ReleaseComObject(excelApp);

            return workbookNames;
        }

        public static bool CloseWorkbookByIdAndName(Dictionary<int, List<string>> pId2WbNames)
        {
            Excel.Application excelApp = null;

            foreach (var item in pId2WbNames)
            {
                try
                {
                    Process process = Process.GetProcessById(item.Key);
                    excelApp = ExcelExtensions.ExcelAppCollection.InnerFromProcess(process);
                }
                catch (Exception ex)
                {
                    // Handle exception if Excel is not running or no workbooks are open
                    Console.WriteLine(ex.Message);
                    return false;
                }

                foreach (Excel.Workbook workbook in excelApp.Workbooks)
                {
                    if (item.Value.Contains(workbook.Name))
                    {
                        workbook.Close(false);
                    }
                }

                if (excelApp.Workbooks.Count == 0) 
                    excelApp.Quit();

                Marshal.ReleaseComObject(excelApp);
            }

            return true;
        }

        public static bool CloseWorkbookByName(List<string> wbNames)
        {
            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                // Handle exception if Excel is not running or no workbooks are open
                Console.WriteLine(ex.Message);
                return false;
            }

            foreach (Excel.Workbook workbook in excelApp.Workbooks)
            {
                if (wbNames.Contains(workbook.Name))
                {
                    workbook.Close(false);
                }
            }

            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return true;
        }

        public static void GenerateTreeView(TreeView tv, Dictionary<int, List<string>> processId2WbNames)
        {
            // Create the root node
            TreeNode rootNode = new TreeNode("All");

            foreach (var item in processId2WbNames)
            {
                TreeNode nProcess = new TreeNode(item.Key.ToString());
                foreach (var item2 in item.Value)
                {
                    nProcess.Nodes.Add(new TreeNode(item2));
                }
                rootNode.Nodes.Add(nProcess);
            }

            // Add the root node to the TreeView
            tv.Nodes.Clear();
            tv.Nodes.Add(rootNode);
            tv.ExpandAll();
        }

        public static Dictionary<int, List<Tuple<int,string>>> GetCheckedNodes(TreeView tv)
        {
            Dictionary<int, List<Tuple<int, string>>> checkedNodes = new Dictionary<int, List<Tuple<int, string>>>();
            TreeNode root = tv.Nodes[0];
            foreach (TreeNode node in root.Nodes)
            {
                List<Tuple<int, string>> wbs = new List<Tuple<int, string>>();
                for (int i = 0; i<node.Nodes.Count; ++i) 
                {
                    if (node.Nodes[i].Checked)
                    {
                        wbs.Add(new Tuple<int, string>(i, node.Nodes[i].Text));
                    }
                }
                if (wbs.Count > 0)
                {
                    checkedNodes[Int32.Parse(node.Text)] = wbs;
                }  
            }
            return checkedNodes;
        }

        public static void UpdateAllDescendants(TreeNode node, bool checkStatus = true)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = checkStatus;
                UpdateAllDescendants(child, checkStatus);
            }
        }

        public static void UncheckAllAscendants(TreeNode node)
        {
            if (node.Parent != null)
            {
                node.Parent.Checked = false;
                UncheckAllAscendants(node.Parent);
            }
        }

        public static void UpdateCheckStatus(TreeNode node, bool checkStatus = true)
        {
            // if unchecked, update all its ascendants to be unchecked
            if (!checkStatus)
            {
                UncheckAllAscendants(node);
            }

            // update all descendats
            node.Checked = checkStatus;
            UpdateAllDescendants(node, checkStatus);
        }
    }
}

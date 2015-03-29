using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Collections;


namespace SpoolDriver
{
    public class WindowsDriver
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }

        /// <summary>
        /// The OpenPrinter function retrieves a handle to the specified printer or print server or other types of handles in the print subsystem.
        /// </summary>
        /// <param name="szPrinter">
        /// A pointer to a null-terminated string that specifies the name of the printer or print server, the printer object, the XcvMonitor, or the XcvPort.
        /// For a printer object use: PrinterName, Job xxxx. For an XcvMonitor, use: ServerName, XcvMonitor MonitorName. For an XcvPort, use: ServerName, XcvPort PortName.
        /// If NULL, it indicates the local printer server.</param>
        /// <param name="hPrinter">A pointer to a variable that receives a handle (not thread safe) to the open printer or print server object.
        /// The phPrinter parameter can return an Xcv handle for use with the XcvData function. For more information about XcvData, see the DDK.</param>
        /// <param name="pd">A pointer to a PRINTER_DEFAULTS structure. This value can be NULL.</param>
        /// <returns>If the function succeeds, the return value is a nonzero value.
        /// If the function fails, the return value is zero.</returns>
        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        /// <summary>
        /// The ClosePrinter function closes the specified printer object.
        /// </summary>
        /// <param name="hPrinter">A handle to the printer object to be closed. This handle is returned by the OpenPrinter or AddPrinter function.</param>
        /// <returns>If the function succeeds, the return value is a nonzero value.
        /// If the function fails, the return value is zero.</returns>
        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);
               

        /// <summary>
        /// The WritePrinter function notifies the print spooler that data should be written to the specified printer.
        /// </summary>
        /// <param name="hPrinter">A handle to the printer. 
        /// Use the OpenPrinter or AddPrinter function to retrieve a printer handle.</param>
        /// <param name="pBytes">A pointer to an array of bytes 
        /// that contains the data that should be written to the printer.</param>
        /// <param name="dwCount">The size, in bytes, of the array.</param>
        /// <param name="dwWritten">A pointer to a value that receives the number of bytes of 
        /// data that were written to the printer.</param>
        /// <returns>If the function succeeds, the return value is a nonzero value. 
        /// If the function fails, the return value is zero.</returns>
        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        [DllImport("winspool.Drv", EntryPoint = "GetSpoolFileHandle", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetSpoolFileHandle(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "CloseSpoolFileHandle", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool CloseSpoolFileHandle(IntPtr hPrinter, IntPtr hSpoolFile);

        [DllImport("Spoolss.dll", EntryPoint = "AbortPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool AbortPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "FlushPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool FlushPrinter(IntPtr hPrinter, IntPtr pBuf, Int32 cbBuf, out Int32 pcWritten, Int32 cSleep);
        
        [DllImport("winspool.Drv", EntryPoint = "PrinterProperties", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool PrinterProperties(IntPtr hWindows, IntPtr hPrinter);


        /// <summary>
        /// Send array of byte to printer
        /// </summary>
        /// <param name="szPrinterName">Printer name (from windows control panel)</param>
        /// <param name="pBytes">pointer to array of byte</param>
        /// <param name="dwCount">array size</param>        
        /// <param name="docName">header visible in spooler print list window</param>
        /// <returns>true if success, otherwise false</returns>
        public static bool SendBytes(string szPrinterName, IntPtr pBytes, Int32 dwCount, string docName)
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false;
            di.pDocName = docName;
            di.pDataType = "RAW";

            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {                
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    if (StartPagePrinter(hPrinter))
                    {
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);                        
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }                
                ClosePrinter(hPrinter);
            }

            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        /// <summary>
        /// Send file to the printer
        /// </summary>
        /// <param name="szPrinterName">Printer name (from windows control panel)</param>
        /// <param name="szFileName">file path wich will be printed</param>
        /// <returns>true if success, otherwise false</returns>
        public static bool SendFile(string szPrinterName, string szFileName)
        {

            FileStream fs = new FileStream(szFileName, FileMode.Open);

            BinaryReader br = new BinaryReader(fs);

            Byte[] bytes = new Byte[fs.Length];
            bool bSuccess = false;

            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);

            bytes = br.ReadBytes(nLength);

            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);

            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);

            bSuccess = SendBytes(szPrinterName, pUnmanagedBytes, nLength, szFileName);

            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        /// <summary>
        /// Send command to the printer. It could be EPL command
        /// </summary>
        /// <param name="szPrinterName">Printer name (from windows control panel)</param>
        /// <param name="szString">Command (EPL)</param>
        /// <param name="docName">Header visible in spooler print list window</param>
        /// <returns>true</returns>
        public static bool SendString(string szPrinterName, string szString, string docName)
        {
            IntPtr pBytes;
            Int32 dwCount;

            dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize;

            pBytes = Marshal.StringToCoTaskMemAnsi(szString);

            SendBytes(szPrinterName, pBytes, dwCount, docName);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }
    }

    public class PrinterHelper
    {
        private IntPtr hPrinter = new IntPtr(0);
        private IntPtr hSpoolFile = new IntPtr(0);
        private string jobName;
        private string printerName;
        private bool withMonitor;
        private Int32 lastError;
        private Int32 lastWritten;        

        #region

        public static event PrintingEventsHandler PrinterOpened;
        //public static event PrintingEventsHandler PrinterClosed;
        //public static event PrintingEventsHandler PrintingStarted;

        #endregion

        #region Properties

        public IntPtr PrinterHandler
        {
            get
            {
                return hPrinter;
            }
        }
        public IntPtr SpoolFileHandler
        {
            get
            {
                return hSpoolFile;
            }
        }
        public string PrinterName
        {
            get
            {
                return printerName;
            }
            set
            {
                printerName = value;
            }
        }
        public string JobName
        {
            get
            {
                return jobName;
            }
            set
            {
                jobName = value;
            }
        }
        public bool ShowMonitor
        {
            set
            {
                withMonitor = value;
            }
            get
            {
                return withMonitor;
            }
        }
        public Int32 LastError
        {
            get
            {
                return this.lastError;
            }
        }
        public Int32 LastWritten
        {
            get
            {
                return this.lastWritten;
            }
        }

        #endregion

        #region Static Method

        public static void PrinterHelperProcParam(object _param)
        {
            System.Collections.ArrayList _arr = (System.Collections.ArrayList)_param;
            if (_arr != null)
            {
                if (_arr.Count >= 4)
                {
                    string printeName = (string)_arr[0];
                    string command = (string)_arr[1];
                    string jobName = (string)_arr[2];
                    bool showMon = (bool)_arr[3];

                    PrinterHelper helper = new PrinterHelper(printeName, jobName, showMon);                    
                    helper.WriteString(command);
                }
            }
        }        
        
        public static void PrinterPropertiesProcParam(object _param)
        {
            if (_param != null)
            {
                Type type = _param.GetType();
                if (type.Equals(typeof(PrinterHelper)))
                {
                    PrinterHelper helper = (PrinterHelper)_param;
                    helper.PrinterProperties();
                }
            }
            
        }

        //public static void PrinterMonitorProcParam(object _param)
        //{
        //    throw new NotImplementedException(Properties.Resources.MsgNotImplemented);    
        //}       

        /// <summary>
        /// Wykonuje komendę w oddzielnym wątku
        /// </summary>
        /// <param name="_printerName">Nazwa drukarki</param>
        /// <param name="_command">Komenda</param>
        /// <param name="_jobName">Nazwa zadania</param>
        /// <param name="_showMonitor">Wyswietla formularz monitorujace (nie zaimplementowane)</param>
        public static void ExecuteCommand(string _printerName, string _command, string _jobName, bool _showMonitor)
        {
            
            System.Collections.ArrayList _params = new System.Collections.ArrayList();
            _params.Add(_printerName);
            _params.Add(_command);
            _params.Add(_jobName);
            _params.Add(_showMonitor);

            Thread printerProc = new Thread(new ParameterizedThreadStart(PrinterHelperProcParam));
            printerProc.Name = String.Format("{0}", _jobName);
            printerProc.Start(_params);           
        }
        
        /// <summary>
        /// Uruchamia formularz parametrów drukarki w odzielnym wątku
        /// </summary>
        /// <param name="_printerName">Nazwa drukarki</param>
        public static void ShowProperties(string _printerName)
        {
            PrinterHelper helper = new PrinterHelper(_printerName);
            Thread thd = new Thread(new ParameterizedThreadStart(PrinterPropertiesProcParam));
            thd.Name = String.Format("{0}", _printerName);
            thd.Start(helper);            
        }    

        #endregion

        public PrinterHelper()
        {
        }
        public PrinterHelper(string _printerName, string _jobName, bool _showMonitor)
        {
            printerName = _printerName;
            jobName = _jobName;
            withMonitor = _showMonitor;
        }
        public PrinterHelper(string _printerName)
        {
            printerName = _printerName;
        }

        public bool WriteBytes(byte[] bytes)
        {
            IntPtr pBytes = Marshal.AllocHGlobal(bytes.Length);
            Marshal.Copy(bytes, 0, pBytes, bytes.Length);
            bool ret = false;
            ret = WriteBytes(pBytes, bytes.Length);
            Marshal.FreeHGlobal(pBytes);
            return ret;
        }
        public bool WriteBytes(IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
            //hPrinter = new IntPtr(0);
            WindowsDriver.DOCINFOA di = new WindowsDriver.DOCINFOA();
            bool bSuccess = false;
            di.pDocName = jobName;
            di.pDataType = "RAW";
                        
            if (this.OpenPrinter())
            {
                
                //if (withMonitor)
                //{
                //    monitorForm = new MonitorForm(this);
                //    thMonitor = new Thread(new ParameterizedThreadStart(PrinterHelper.PrinterMonitorProcParam));                  
                //    thMonitor.Start(this);
                //}
                
                if (WindowsDriver.StartDocPrinter(hPrinter, 1, di))
                {
                    if (WindowsDriver.StartPagePrinter(hPrinter))
                    {
                        bSuccess = WindowsDriver.WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        WindowsDriver.EndPagePrinter(hPrinter);
                    }
                    WindowsDriver.EndDocPrinter(hPrinter);
                }

                //if (withMonitor)
                //{
                //    monitorForm.Close();
                //    //thMonitor.Abort();
                //}
                this.ClosePrinter();
            }

            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            lastError = dwError;
            lastWritten = dwWritten;
            return bSuccess;
        }

        public bool WriteString(string _cmd)
        {
            IntPtr pBytes;
            Int32 dwCount;            
            dwCount = (_cmd.Length + 1) * Marshal.SystemMaxDBCSCharSize;

            pBytes = Marshal.StringToCoTaskMemAnsi(_cmd);

            this.WriteBytes(pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }

        public bool WriteFile(string _fileName)
        {

            FileStream fs = new FileStream(_fileName, FileMode.Open);

            BinaryReader br = new BinaryReader(fs);

            Byte[] bytes = new Byte[fs.Length];
            bool bSuccess = false;

            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);

            bytes = br.ReadBytes(nLength);

            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);

            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);

            bSuccess = WriteBytes(pUnmanagedBytes, nLength);

            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        public bool OpenPrinter()
        {
            if (WindowsDriver.OpenPrinter(printerName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                this.raisePrinterOpened();
                return true;
            }
            return false;
        }
        public bool ClosePrinter()
        {
            return WindowsDriver.ClosePrinter(hPrinter);
        }

        public bool PrinterProperties()
        {
            bool result = false;
            if (this.OpenPrinter())
            {
                result = WindowsDriver.PrinterProperties(IntPtr.Zero, hPrinter);
            }
            this.ClosePrinter();
            return result;
        }

        #region Event Triggers

        void raisePrinterOpened()
        {
            if (PrinterOpened != null)
            {
                PrinterOpened(this);
            }
        }
        
        #endregion

    }
    
    public delegate void PrintingEventsHandler(object e);

}

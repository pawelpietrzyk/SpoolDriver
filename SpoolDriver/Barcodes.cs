using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Ports;
using System.Globalization;
using SpoolDriver;

namespace SpoolDriver.Barcodes.EPL
{
    public class BarcodeDesign
    {
        System.Collections.ArrayList m_elements;
        StringBuilder m_cmd;
        string m_spoolerHead;
        int m_copies;

        public string Command
        {
            get
            {
                return this.createCommand();
            }
            set
            {
                m_cmd = new StringBuilder(value);
            }
        }
        public string SpoolerHeader
        {
            get
            {
                return m_spoolerHead;
            }
            set
            {
                m_spoolerHead = value;
            }
        }
        public int Copies
        {
            get
            {
                return m_copies;
            }
            set
            {
                m_copies = value;
            }
        }

        public BarcodeDesign()
        {
            m_elements = new System.Collections.ArrayList();
            m_spoolerHead = "Barcode";
            m_copies = 1;
        }

        public void AddText(BarcodeTextDef _barTextObj)
        {
            m_elements.Add(_barTextObj.Command);
        }
        public void AddText(string _p1, string _p2, string _p3, string _p4, string _p5, string _p6, string _p7, string _text)
        {
            BarcodeTextDef _barTextObj = new BarcodeTextDef(_p1, _p2, _p3, _p4, _p5, _p6, _p7, _text);
            m_elements.Add(_barTextObj.Command);
        }
        public void AddBarcode(BarcodeDef _barDef)
        {
            m_elements.Add(_barDef.Command);
        }
        public void AddBarcode(string _p1, string _p2, string _p3, string _p4, string _p5, string _p6, string _p7, string _p8, string _data)
        {
            BarcodeDef _barDef = new BarcodeDef(_p1, _p2, _p3, _p4, _p5, _p6, _p7, _p8, _data);
            m_elements.Add(_barDef.Command);
        }

        public void ExecuteDesign(string printerName, bool showMonitor)
        {
            if (!String.IsNullOrEmpty(printerName))
            {
                PrinterHelper.ExecuteCommand(printerName, this.Command, this.SpoolerHeader, showMonitor);
            }

        }

        private string createCommand()
        {
            m_cmd = new StringBuilder();
            m_cmd.AppendLine("N");
            foreach (string element in m_elements)
            {
                m_cmd.Append(element);
            }
            //m_cmd.AppendLine("\n");
            m_cmd.AppendLine(String.Format("P{0}", this.Copies));
            return m_cmd.ToString();            
        }
        private int CDot(int _valueMM)
        {
            return (_valueMM * 8);
        }
    }

    public class BarcodeDef
    {
        #region LocalVariables
        string m_P1;
        string m_P2;
        string m_P3;
        string m_P4;
        string m_P5;
        string m_P6;
        string m_P7;
        string m_P8;
        string m_DATA;
        BarcodeTypes m_barcodeType;
        string m_cmd;
        #endregion

        #region Properties

        public string P1
        {
            get
            {
                return m_P1;
            }
            set
            {
                m_P1 = value;
            }
        }
        public string P2
        {
            get
            {
                return m_P2;
            }
            set
            {
                m_P2 = value;
            }
        }
        public string P3
        {
            get
            {
                return m_P3;
            }
            set
            {
                m_P3 = value;
            }
        }
        public string P4
        {
            get
            {
                return m_P4;
            }
            set
            {
                m_P4 = value;
            }
        }
        public string P5
        {
            get
            {
                return m_P5;
            }
            set
            {
                m_P5 = value;
            }
        }
        public string P6
        {
            get
            {
                return m_P6;
            }
            set
            {
                m_P6 = value;
            }
        }
        public string P7
        {
            get
            {
                return m_P7;
            }
            set
            {
                m_P7 = value;
            }
        }
        public string P8
        {
            get
            {
                return m_P8;
            }
            set
            {
                m_P8 = value;
            }
        }
        public string DATA
        {
            get
            {
                return m_DATA;
            }
            set
            {
                m_DATA = value;
            }
        }
        public BarcodeTypes BarcodeType
        {
            get
            {
                return this.m_barcodeType;
            }
            set
            {
                this.m_barcodeType = value;
            }
        }

        public string Command
        {
            get
            {
                this.m_cmd = this.createCommand();
                return this.m_cmd.ToString();
            }

        }

        #endregion


        
        public BarcodeDef(string _p4, string _p5, string _p6, BarcodeTypes _type)
        {
            this.P4 = _p4;
            this.P5 = _p5;
            this.P6 = _p6;
            this.BarcodeType = _type;
        }

        public BarcodeDef(string _p1, string _p2, string _p3, string _p4, string _p5, string _p6, string _p7, string _p8, string _data)
        {
            this.P1 = _p1;
            this.P2 = _p2;
            this.P3 = _p3;
            this.P4 = _p4;
            this.P5 = _p5;
            this.P6 = _p6;
            this.P7 = _p7;
            this.P8 = _p8;
            this.DATA = _data;
        }

        public static BarcodeDef getBarcodeTypeDef(BarcodeTypes _type, int _scale)
        {
            int scale_min;
            int scale_max;            
            int scale;

            switch (_type)
            {
                case BarcodeTypes.Code39_Std:
                    scale_min = 1;
                    scale_max = 10;
                    scale = (int)Decimal.Round(((decimal)_scale / (decimal)100) * (scale_max - scale_min), 0);
                    return new BarcodeDef("3", scale.ToString(), "7", _type);
                default:
                    scale_min = 1;
                    scale_max = 10;
                    scale = (int)Decimal.Round(((decimal)_scale / (decimal)100) * (scale_max - scale_min), 0);
                    return new BarcodeDef("3", scale.ToString(), "7", _type);
            }
        }

        private string createCommand()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format(
                CultureInfo.InvariantCulture,
                "B{0},{1},{2},{3},{4},{5},{6},{7},\"{8}\"",
                new object[] { P1, P2, P3, P4, P5, P6, P7, P8, DATA }));
            return sb.ToString();
        }

    }

    public class BarcodeTextDef
    {
        #region LocalVariables
        string m_P1;
        string m_P2;
        string m_P3;
        string m_P4;
        string m_P5;
        string m_P6;
        string m_P7;
        string m_P8;
        string m_TEXT;        
        string m_cmd;
        #endregion

        #region Properties

        public string P1
        {
            get
            {
                return m_P1;
            }
            set
            {
                m_P1 = value;
            }
        }
        public string P2
        {
            get
            {
                return m_P2;
            }
            set
            {
                m_P2 = value;
            }
        }
        public string P3
        {
            get
            {
                return m_P3;
            }
            set
            {
                m_P3 = value;
            }
        }
        public string P4
        {
            get
            {
                return m_P4;
            }
            set
            {
                m_P4 = value;
            }
        }
        public string P5
        {
            get
            {
                return m_P5;
            }
            set
            {
                m_P5 = value;
            }
        }
        public string P6
        {
            get
            {
                return m_P6;
            }
            set
            {
                m_P6 = value;
            }
        }
        public string P7
        {
            get
            {
                return m_P7;
            }
            set
            {
                m_P7 = value;
            }
        }
        public string P8
        {
            get
            {
                return m_P8;
            }
            set
            {
                m_P8 = value;
            }
        }
        public string TEXT
        {
            get
            {
                return m_TEXT;
            }
            set
            {
                m_TEXT = value;
            }
        }
        
        public string Command
        {
            get
            {
                this.m_cmd = this.createCommand();
                return this.m_cmd.ToString();
            }

        }

        #endregion

        /// <summary>
        /// ASCII Text
        /// </summary>
        /// <param name="_p1">Horizontal start positionX) in dots</param>
        /// <param name="_p2">Vertical start position (Y) in dots</param>
        /// <param name="_p3">Rotation:
        /// 0 = normal (no rotation),
        /// 1 = 90 degress,
        /// 2 = 180 degress,
        /// 3 = 270 degeress</param>
        /// <param name="_p4">Font selection:
        /// 1 = [203dpi: 20.3 cpi, 6 pts, (8x12 dots)], [300 dpi: 25 cpi, 4 pts, (12x20 dots)];
        /// 2 = [203dpi: 16.9 cpi, 7 pts, (10x16 dots)], [300 dpi: 18.75 cpi, 6 pts, (16x28 dots)];
        /// 3 = [203dpi: 14.5 cpi, 10 pts, (12x20 dots)], [300 dpi: 15 cpi, 8 pts, (20x36 dots)];
        /// 4 = [203dpi: 12.7 cpi, 12 pts, (14x24 dots)], [300 dpi: 12.5 cpi, 10 pts, (24x44 dots)];
        /// 5 = [203dpi: 5.6 cpi, 24 pts, (32x48 dots)], [300 dpi: 6.25 cpi, 21 pts, (48x80 dots)];
        /// 6 = [203dpi: Numeric Only (14 x 19 dots)], [300 dpi: Numeric Only (14 x 19 dots)];
        /// 7 = [203dpi: Numeric Only (14 x 19 dots)], [300 dpi: Numeric Only (14 x 19 dots)]</param>
        /// <param name="_p5">Horizontal multipler, Accepted values: 1-6, 8</param>
        /// <param name="_p6">Vertical multipler, Accepted values: 1-9</param>
        /// <param name="_p7">Reverse image. Accepted values: N = normal, R = reverse image</param>
        /// <param name="_text">Data field</param>
        public BarcodeTextDef(string _p1, string _p2, string _p3, string _p4, string _p5, string _p6, string _p7, string _text)
        {
            this.P1 = _p1;
            this.P2 = _p2;
            this.P3 = _p3;
            this.P4 = _p4;
            this.P5 = _p5;
            this.P6 = _p6;
            this.P7 = _p7;
            this.TEXT = _text;            
        }

        public BarcodeTextDef(int _posX, int _posY, int _fontSize, string _text)
        {
            this.P1 = _posX.ToString();
            this.P2 = _posY.ToString();
            this.P3 = "0";
            if (_fontSize > 7)
            {
                _fontSize = 7;
            }
            if (_fontSize < 1)
            {
                _fontSize = 1;
            }
            this.P4 = _fontSize.ToString();
            this.P5 = "1";
            this.P6 = "1";
            this.P7 = "N";
            this.TEXT = _text;
        }
        
        private string createCommand()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format(
                CultureInfo.InvariantCulture,
                "A{0},{1},{2},{3},{4},{5},{6},\"{7}\"",
                new object[] { P1, P2, P3, P4, P5, P6, P7, TEXT }));
            return sb.ToString();
        }

    }

    public enum BarcodeTypes
    {
        Code39_Std,
        Code39_CheckDigit,
        Code128_UCC,
        Code128_ABC,
        Code128_A,
        Code128_B,
        Code128_C
    }
    public enum Rotation
    {
        NoRotation = 0,
        Degress90 = 1,
        Degress180 = 2,
        Degress270 = 3
    }



}

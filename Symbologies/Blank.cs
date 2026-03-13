using System;
using System.Collections.Generic;
using System.Text;
using BatDongSan.Models.BarcodeLib;

namespace BatDongSan.Symbologies
{
    /// <summary>
    ///  Blank encoding template
    ///  Written by: Brad Barnhill
    /// </summary>
    class Blank: BarcodeCommon, IBarcode
    {
        
        #region IBarcode Members

        public string Encoded_Value
        {
            get { throw new NotImplementedException(); }
        }

        #endregion
    }
}

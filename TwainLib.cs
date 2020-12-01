//#define twain64

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TwainLib
{
    public class Twain
    {
        #region Declarations
        #region COM Declarations
#if twain64
        // ------ DSM entry point DAT_ variants:
        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSMparent([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, ref IntPtr refptr);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSMident([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwIdentity idds);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSMstatus([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwStatus dsmstat);


        // ------ DSM entry point DAT_ variants to DS:
        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSuserif([In, Out] TwIdentity origin, [In, Out] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, TwUserInterface guif);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSevent([In, Out] TwIdentity origin, [In, Out] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, ref TwEvent evt);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSstatus([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwStatus dsmstat);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DScap([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwCapability capa);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSiinf([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwImageInfo imginf);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSixfer([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, ref IntPtr hbitmap);

        [DllImport("twaindsm.dll", EntryPoint = "#1")]
        private static extern TwRC DSpxfer([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwPendingXfers pxfr);
#else
        // ------ DSM entry point DAT_ variants:
        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSMparent([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, ref IntPtr refptr);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSMident([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwIdentity idds);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSMstatus([In, Out] TwIdentity origin, IntPtr zeroptr, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwStatus dsmstat);


        // ------ DSM entry point DAT_ variants to DS:
        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSuserif([In, Out] TwIdentity origin, [In, Out] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, TwUserInterface guif);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSevent([In, Out] TwIdentity origin, [In, Out] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, ref TwEvent evt);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSstatus([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwStatus dsmstat);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DScap([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwCapability capa);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSiinf([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwImageInfo imginf);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSixfer([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, ref IntPtr hbitmap);

        [DllImport("twain_32.dll", EntryPoint = "#1")]
        private static extern TwRC DSpxfer([In, Out] TwIdentity origin, [In] TwIdentity dest, TwDG dg, TwDAT dat, TwMSG msg, [In, Out] TwPendingXfers pxfr);
#endif

        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GlobalAlloc(int flags, int size);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GlobalLock(IntPtr handle);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern bool GlobalUnlock(IntPtr handle);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GlobalFree(IntPtr handle);

        [DllImport("user32.dll", ExactSpelling = true)]
        private static extern int GetMessagePos();
        [DllImport("user32.dll", ExactSpelling = true)]
        private static extern int GetMessageTime();


        [DllImport("gdi32.dll", ExactSpelling = true)]
        private static extern int GetDeviceCaps(IntPtr hDC, int nIndex);

        [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr CreateDC(string szdriver, string szdevice, string szoutput, IntPtr devmode);

        [DllImport("gdi32.dll", ExactSpelling = true)]
        private static extern bool DeleteDC(IntPtr hdc);


        [DllImport("gdiplus.dll", ExactSpelling = true)]
        private static extern int GdipCreateBitmapFromGdiDib(IntPtr bminfo, IntPtr pixdat, ref IntPtr image);
        #endregion COM Declarations

        private const short CountryUSA = 1;
        private const short LanguageUSA = 13;

        internal IntPtr hwnd;
        private TwIdentity appid;
        private TwIdentity srcds;
        private TwEvent evtmsg;
        private WINMSG winmsg;

        public enum TwainCommand
        {
            Not = -1,
            Null = 0,
            TransferReady = 1,
            CloseRequest = 2,
            CloseOk = 3,
            DeviceEvent = 4
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        internal struct WINMSG
        {
            public IntPtr hwnd;
            public int message;
            public IntPtr wParam;
            public IntPtr lParam;
            public int time;
            public int x;
            public int y;
        }


        [StructLayout(LayoutKind.Sequential, Pack = 2)]
        internal class BITMAPINFOHEADER
        {
            public int biSize = 0;
            public int biWidth = 0;
            public int biHeight = 0;
            public short biPlanes = 0;
            public short biBitCount = 0;
            public int biCompression = 0;
            public int biSizeImage = 0;
            public int biXPelsPerMeter = 0;
            public int biYPelsPerMeter = 0;
            public int biClrUsed = 0;
            public int biClrImportant = 0;
        }
        #endregion Declarations

        #region Constructors
        public Twain()
        {
            appid = new TwIdentity();
            appid.Id = 0;
            appid.Version.MajorNum = 1;
            appid.Version.MinorNum = 1;
            appid.Version.Language = LanguageUSA;
            appid.Version.Country = CountryUSA;
            appid.Version.Info = "Hack 1";
            appid.ProtocolMajor = TwProtocol.Major;
            appid.ProtocolMinor = TwProtocol.Minor;
            appid.SupportedGroups = (int)(TwDG.Image | TwDG.Control);
            appid.Manufacturer = "NETMaster";
            appid.ProductFamily = "Freeware";
            appid.ProductName = "Hack";

            srcds = new TwIdentity();
            srcds.Id = 0;

            evtmsg.EventPtr = Marshal.AllocHGlobal(Marshal.SizeOf(winmsg));
        }

        ~Twain()
        {
            Marshal.FreeHGlobal(evtmsg.EventPtr);
        }
        #endregion Constructors

        #region Properties
        public static int ScreenBitDepth
        {
            get
            {
                IntPtr screenDC = CreateDC("DISPLAY", null, null, IntPtr.Zero);
                int bitDepth = GetDeviceCaps(screenDC, 12);
                bitDepth *= GetDeviceCaps(screenDC, 14);
                DeleteDC(screenDC);
                return bitDepth;
            }
        }
        #endregion Properties

        #region Methods
        /// <summary>
        /// 
        /// </summary>
        /// <param name="hwndp"></param>
        public void Init(IntPtr hwndp)
        {
            Finish();
            TwRC rc = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.OpenDSM, ref hwndp);

            if (rc == TwRC.Success)
            {
                rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetDefault, srcds);
                if (rc == TwRC.Success)
                    hwnd = hwndp;
                else
                    rc = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.CloseDSM, ref hwndp);
            }
        }

        /// <summary>
        /// Opens Twain UI and allows user to choose device from lsit.
        /// </summary>
        public void Select()
        {
            TwRC rc;
            CloseSrc();
            if (appid.Id == 0)
            {
                Init(hwnd);
                if (appid.Id == 0)
                    return;
            }

            rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.UserSelect, srcds);
        }

        /// <summary>
        /// Gets list of device names.
        /// Use with Set.
        /// </summary>
        /// <returns>List(Of String)</returns>
        public List<String> Sources()
        {
            List<String> returnValue = default(List<String>);
            TwRC rc = default(TwRC);

            returnValue = new List<String>();
            rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetFirst, srcds);
            while (rc != TwRC.EndOfList)
            {
                if (rc == TwRC.Success)
                {
                    returnValue.Add(srcds.ProductName);
                }
                else
                {
                    if (returnValue.Count == 0)
                    {
                        //Note: when the list is empty, twain will return Failure code instead of EndOfList
                        break;
                    }
                    else
                    {
                        throw new Exception(String.Format("Unable to retrieve device item, index #{0}: \n{1}", returnValue.Count - 1, rc.ToString()));
                    }
                }

                rc =  DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetNext, srcds);
            };
            
            return returnValue;
        }

        ///// <summary>
        ///// Gets list of device identities.
        ///// Used by Set.
        ///// </summary>
        ///// <returns>List<TwIdentity></returns>
        //internal List<TwIdentity> SourceIdentities()
        //{
        //    List<TwIdentity> returnValue = default(List<TwIdentity>);
        //    TwRC rc = default(TwRC);

        //    returnValue = new List<TwIdentity>();
        //    rc =  DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetFirst, srcds);
        //    while (rc != TwRC.EndOfList)
        //    {
        //        if (rc == TwRC.Success)
        //        {
        //            returnValue.Add(srcds);
        //        }
        //        else
        //        {
        //            throw new Exception(String.Format("Unable to retrieve device item, index #{0}: \n{1}", returnValue.Count - 1, rc.ToString()));
        //        }

        //        rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetNext, srcds);
        //    };
            
        //    return returnValue;
        //}


        ///// <summary>
        ///// Set (select) device by name.
        ///// Neither approach works; set fails and iterating makes no difference
        ///// </summary>
        ///// <param name="sourceName"></param>
        //public void Set(String sourceName)
        //{
        //    TwRC rc = default(TwRC);
        //    #region call set command
        //    //List<TwIdentity> identities = default(List<TwIdentity>);

        //    //identities = SourceIdentities();
        //    //srcds = identities.Find(i => i.ProductName == sourceName);

        //    //if (srcds != null)
        //    //{
        //    //    rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.Set, srcds);
        //    //    if (rc != TwRC.Success)
        //    //    {
        //    //        throw new Exception(String.Format("Unable to set ProductName to {0}: \n{1}", sourceName, rc.ToString()));
        //    //    }
        //    //}
        //    //else
        //    //{ 
        //    //    throw new Exception(String.Format("Unable to find device with ProductName {0}", sourceName));
        //    //}
        //    #endregion call set command
        //    #region iterate devices and leave on match
        //    rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetFirst, srcds);
        //    while (rc != TwRC.EndOfList)
        //    {
        //        if (rc == TwRC.Success)
        //        {
        //            if (srcds.ProductName == sourceName)
        //            {
        //                //found; stop on this item
        //                break;
        //            }
        //        }
        //        else
        //        {
        //            //report error
        //            throw new Exception(String.Format("Unable to locate and set device item to {0}: \n{1}", sourceName, rc.ToString()));
        //        }

        //        rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetNext, srcds);
        //    };
            
        //    //report no matches
        //    if (rc == TwRC.EndOfList)
        //    {
        //        throw new Exception(String.Format("Unable to locate and set device item to {0}: \n{1}", sourceName, rc.ToString()));
        //    }
        //    #endregion iterate devices and leave on match
        //}

        /// <summary>
        /// Opens Twain UI and allows user to select image(s).
        /// </summary>
        public void Acquire()
        {
            TwRC rc;
            CloseSrc();
            if (appid.Id == 0)
            {
                Init(hwnd);
                if (appid.Id == 0)
                    return;
            }
            rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.OpenDS, srcds);
            if (rc != TwRC.Success)
                return;

            TwCapability cap = new TwCapability(TwCap.XferCount, 1);
            rc = DScap(appid, srcds, TwDG.Control, TwDAT.Capability, TwMSG.Set, cap);
            if (rc != TwRC.Success)
            {
                CloseSrc();
                return;
            }

            TwUserInterface guif = new TwUserInterface();
            guif.ShowUI = 1;
            guif.ModalUI = 1;
            guif.ParentHand = hwnd;
            rc = DSuserif(appid, srcds, TwDG.Control, TwDAT.UserInterface, TwMSG.EnableDS, guif);
            if (rc != TwRC.Success)
            {
                CloseSrc();
                return;
            }
        }

        /// <summary>
        /// Performs transfer of image(s) from device.
        /// </summary>
        /// <returns></returns>
        public List<Image> TransferPictures()
        {
            List<Image> returnValue = new List<Image>();
            if (srcds.Id == 0)
                return returnValue;

            TwRC rc;
            IntPtr hbitmap = IntPtr.Zero;
            TwPendingXfers pxfr = new TwPendingXfers();

            do
            {
                pxfr.Count = 0;
                hbitmap = IntPtr.Zero;

                TwImageInfo iinf = new TwImageInfo();
                rc = DSiinf(appid, srcds, TwDG.Image, TwDAT.ImageInfo, TwMSG.Get, iinf);
                if (rc != TwRC.Success)
                {
                    CloseSrc();
                    return returnValue;
                }

                rc = DSixfer(appid, srcds, TwDG.Image, TwDAT.ImageNativeXfer, TwMSG.Get, ref hbitmap);
                if (rc != TwRC.XferDone)
                {
                    CloseSrc();
                    return returnValue;
                }

                rc = DSpxfer(appid, srcds, TwDG.Control, TwDAT.PendingXfers, TwMSG.EndXfer, pxfr);
                if (rc != TwRC.Success)
                {
                    CloseSrc();
                    return returnValue;
                }
                
                returnValue.Add(bitmapFromDIB(hbitmap));
            }
            while (pxfr.Count != 0);

            rc = DSpxfer(appid, srcds, TwDG.Control, TwDAT.PendingXfers, TwMSG.Reset, pxfr);
            return returnValue;
        }

        /// <summary>
        /// Passes WM_ to Twain for translation into TwainCommand.
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        public TwainCommand PassMessage(ref Message m)
        {
            if (srcds.Id == 0)
                return TwainCommand.Not;

            int pos = GetMessagePos();

            winmsg.hwnd = m.HWnd;
            winmsg.message = m.Msg;
            winmsg.wParam = m.WParam;
            winmsg.lParam = m.LParam;
            winmsg.time = GetMessageTime();

            unchecked
            {
                winmsg.x = (short)pos;
                winmsg.y = (short)(pos >> 16);
            }

            Marshal.StructureToPtr(winmsg, evtmsg.EventPtr, false);
            evtmsg.Message = 0;
            TwRC rc = DSevent(appid, srcds, TwDG.Control, TwDAT.Event, TwMSG.ProcessEvent, ref evtmsg);
            if (rc == TwRC.NotDSEvent)
                return TwainCommand.Not;
            if (evtmsg.Message == (short)TwMSG.XFerReady)
                return TwainCommand.TransferReady;
            if (evtmsg.Message == (short)TwMSG.CloseDSReq)
                return TwainCommand.CloseRequest;
            if (evtmsg.Message == (short)TwMSG.CloseDSOK)
                return TwainCommand.CloseOk;
            if (evtmsg.Message == (short)TwMSG.DeviceEvent)
                return TwainCommand.DeviceEvent;

            return TwainCommand.Null;
        }

        /// <summary>
        /// Closes Twain device source
        /// </summary>
        public void CloseSrc()
        {
            TwRC rc;
            if (srcds.Id != 0)
            {
                TwUserInterface guif = new TwUserInterface();
                rc = DSuserif(appid, srcds, TwDG.Control, TwDAT.UserInterface, TwMSG.DisableDS, guif);
                rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.CloseDS, srcds);
            }
        }

        /// <summary>
        /// Closes Twain device source mamager.
        /// </summary>
        public void Finish()
        {
            TwRC rc;
            CloseSrc();
            if (appid.Id != 0)
                rc = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.CloseDSM, ref hwnd);
            appid.Id = 0;
        }

        #region Static Methods
        /// <summary>
        /// converts the TWAIN device result to .NET Bitmap
        /// </summary>
        /// <param name="dibhand">pointer to aquired image</param>
        /// <returns>TWAIN device as .NET Bitmap</returns>
        internal static Bitmap bitmapFromDIB(IntPtr dibhand)
        {
            IntPtr bmpptr = GlobalLock(dibhand);
            IntPtr pixptr = GetPixelInfo(bmpptr);

            IntPtr pBmp = IntPtr.Zero;
            int status = GdipCreateBitmapFromGdiDib(bmpptr, pixptr, ref pBmp);

            if ((status == 0) && (pBmp != IntPtr.Zero))
            {
                MethodInfo mi = typeof(Bitmap).GetMethod("FromGDIplus", BindingFlags.Static | BindingFlags.NonPublic);
                if (mi == null)
                    return null;

                Bitmap result = new Bitmap(mi.Invoke(null, new object[] { pBmp }) as Bitmap);

                GlobalFree(dibhand);
                dibhand = IntPtr.Zero;

                return result;
            }
            else
                return null;
        }

        internal static IntPtr GetPixelInfo(IntPtr bmpptr)
        {
            BITMAPINFOHEADER bmi = new BITMAPINFOHEADER();
            Marshal.PtrToStructure(bmpptr, bmi);

            Rectangle bmprect = new Rectangle(0, 0, bmi.biWidth, bmi.biHeight);

            if (bmi.biSizeImage == 0)
                bmi.biSizeImage = ((((bmi.biWidth * bmi.biBitCount) + 31) & ~31) >> 3) * bmi.biHeight;

            int p = bmi.biClrUsed;
            if ((p == 0) && (bmi.biBitCount <= 8))
                p = 1 << bmi.biBitCount;
            p = (p * 4) + bmi.biSize + (int)bmpptr;
            return (IntPtr)p;
        }
        #endregion Static Methods
        #endregion Methods
    } // class Twain

}

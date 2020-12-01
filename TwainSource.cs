/////////////////////////////////////////////////////////////////////////////////
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
/////////////////////////////////////////////////////////////////////////////////


using System;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace TwainLib
{
    /// 
    /// a simple API for TWAIN device
    /// 
    public class TwainSource : IMessageFilter, IDisposable
    {

        private Twain twain = default(Twain);
        private Boolean messageFilterIsActive = default(Boolean);
        // hack: TwainCommand.Null is sent continuously after TWAIN device is not available or when once preview is done
        private Int32 twainCommandNullCount = 0;

        /// 
        /// all aquired pictures from TWAIN device
        /// 
        public List<Image> ScannedImages = new List<Image>();

        /// 
        /// constructor
        /// 
        /// parent window handle
        public TwainSource(IntPtr hwndp, EventHandler scanFinished)
        {
            //Reset(hwndp, scanFinished);
            try
            {
                twain = new Twain();
                twain.Init(hwndp);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to init Twain", ex);
            }

            try
            {
                if (scanFinished != null)
                {
                    ScanFinished += scanFinished;
                }
                else
                {
                    throw new Exception("ScanFinished handler parameter was null");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to init ScanFinished handler on Twain", ex);
            }
        }

        /// 
        /// clear ressources
        /// 
        public void Dispose()
        {
            if (TransferStarted != null)
            {
                TransferStarted = null;
            }
            if (TransferFinished != null)
            {
                TransferFinished = null;
            }
            if (ScanStarted != null)
            {
                ScanStarted = null;
            }
            if (ScanFinished != null)
            {
                ScanFinished = null;
            }
            twain.Finish();
        }

        /// 
        /// TWAIN device begins to transfer acquired images
        /// 
        public event EventHandler TransferStarted;

        /// 
        /// TWAIN device finished transferring acquired images
        /// 
        public event EventHandler TransferFinished;

        /// 
        /// TWAIN device begins to acquire images
        /// 
        public event EventHandler ScanStarted;

        /// 
        /// TWAIN device finished acquiring images
        /// 
        public event EventHandler ScanFinished;
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<String> GetSources()
        {
            List<String> returnValue = default(List<String>);

            returnValue = twain.Sources();

            return returnValue;
        }
        
        ///// <summary>
        ///// Set Twain device to the one with the specified name.
        ///// </summary>
        ///// <param name="sourceName"></param>
        //public void SetSource(String sourceName)
        //{
        //    twain.Set(sourceName);
        //}

        /// <summary>
        /// Display Twain UI for choosing a TWAIN device.
        /// </summary>
        public void SelectSource()
        {
            twain.Select();
        }

        /// <summary>
        /// Direct call to Acquire without messaging will not 
        /// trigger or respond to Transfer Ready;
        /// requires separate call to TransferPictures
        /// </summary>
        public void Acquire()
        {
            //scan UI will appear here; after selecting Scan transfer will be triggered
            twain.Acquire();
            
            //trigger scan started
            if (ScanStarted != null)
            {
                ScanStarted(this, new EventArgs());
            }
        }

        /// <summary>
        /// Direct call to TransferPictures;
        /// required if messaging not used with Acquire.
        /// </summary>
        public void Transfer()
        {
            if (TransferStarted != null)
            {
                TransferStarted(this, new EventArgs());
            }

            //get pics from twain via xfer
            List<Image> images = twain.TransferPictures();

            //append pics to scanned image list
            foreach (Image image in images)
            {
                ScannedImages.Add(image);
            }

            EndingScan();
            twain.CloseSrc();

            if (TransferFinished != null)
            {
                TransferFinished(this, new EventArgs());
            }
        }

        /// <summary>
        /// Displays Twain UI to acquire (select) and transfer images.
        /// Call to Acquire using messaging; will trigger call to TransferPictures.
        /// </summary>
        public void AcquireAndTransfer()
        {
            if (!messageFilterIsActive)
            {
                messageFilterIsActive = true;
                Application.AddMessageFilter(this);

                //trigger scan started
                if (ScanStarted != null)
                {
                    ScanStarted(this, new EventArgs());
                }
            }
            twain.Acquire();
        }

        /// <summary>
        /// stop listening for TWAIN device results
        /// </summary>
        public void EndingScan()
        {
            if (messageFilterIsActive)
            {
                Application.RemoveMessageFilter(this);
                messageFilterIsActive = false;

                if (ScanFinished != null)
                {
                    ScanFinished(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// receives messages and handles TWAIN device messages
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        Boolean IMessageFilter.PreFilterMessage(ref Message message)
        {
            Boolean returnValue = default(Boolean);
            Twain.TwainCommand twainCommand = twain.PassMessage(ref message);

            switch (twainCommand)
            {
                case Twain.TwainCommand.Not:
                    {
                        //return default (false)
                        break;
                    }
                case Twain.TwainCommand.CloseRequest:
                    {
                        EndingScan();
                        twain.CloseSrc();

                        returnValue = true;
                        break;
                    }
                case Twain.TwainCommand.CloseOk:
                    {
                        EndingScan();
                        twain.CloseSrc();

                        returnValue = true;
                        break;
                    }
                case Twain.TwainCommand.DeviceEvent:
                    {
                        returnValue = true;
                        break;
                    }
                case Twain.TwainCommand.TransferReady:
                    {
                        if (TransferStarted != null)
                        {
                            TransferStarted(this, new EventArgs());
                        }

                        //get pics from twain via xfer (scanner may begin to perform scan here)
                        List<Image> images = twain.TransferPictures();

                        //append pics to scanned image list
                        foreach (Image image in images)
                        {
                            ScannedImages.Add(image);
                        }

                        EndingScan();//TODO:DEBUG:twain.ScanFinished delegate being cleared here
                        twain.CloseSrc();

                        if (TransferFinished != null)
                        {
                            TransferFinished(this, new EventArgs());
                        }

                        returnValue = true;
                        break;
                    }
                case Twain.TwainCommand.Null:
                    {
                        twainCommandNullCount++;

                        if (twainCommandNullCount > 25)
                        {
                            twainCommandNullCount = 0;
                            
                            EndingScan();
                            twain.CloseSrc();
                        }

                        returnValue = true;
                        break;
                    }
            }

            return returnValue;
        }
    }
}


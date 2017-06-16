using System;
using Microsoft.VisualBasic.CompilerServices;
using System.Runtime.InteropServices;
using System.Collections;
using System.Windows.Forms;
using System.Drawing;
using stdole;
using System.Drawing.Imaging;
using System.Runtime.CompilerServices;

namespace UpgradeHelpers.SupportHelper
{

    /// <summary>
    /// Enumerations used by showForm
    /// </summary>
    public enum FormShowConstants
    {
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbModal.</summary>
        Modal,

        /// <summary>Equivalent to the Visual Basic 6.0 constant vbModeless..</summary>
        Modeless
    }

    /// <summary>
    /// Enumerations used by printing Pixels and Twips
    /// </summary>
    public enum ScaleMode
    {
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbPoints.</summary>
        Points = 2,
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbCharacters.</summary>
        Characters = 4,
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbInches.</summary>
        Inches,
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbMillimeters.</summary>
        Millimeters,
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbCentimeters.</summary>
        Centimeters,
        /// <summary>Equivalent to the Visual Basic 6.0 constant vbHimetric.</summary>
        Himetric
    }

    /// <summary>
    /// This is a replacement class of Microsoft.VisualBasic.Compatibility.Support class
    /// </summary>
    public class Support
    {
        /* 
         * UNUSED CONSTANTS
                private const int TwipsPerCharHoriz = 120;
                private const int TwipsPerCharVert = 240;
                private const double TwipsPerCm = 566.92913385826773;
                private const double TwipsPerHiMetric = 0.56692913385826771;
                private const int TwipsPerInch = 1440;
                private const double TwipsPerMm = 56.692913385826778;
                private const int TwipsPerPoint = 20;
                private const double CmPerInch = 2.54;
                private const double HiMetricPerInch = 2540.0;
                private const int HimetricPerMm = 100;
        */
        private static bool _isTwipsPerPixelSetUp;

        private static double _twipsPerPixelX;
        private static double _twipsPerPixelY;

        private static Hashtable _resManagers = new Hashtable();

        static Support()
        {
            // Note: this type is marked as 'beforefieldinit'.
            MResManagers = new Hashtable();
        }

        /// <summary>
        /// The seter and geter of _mResManagers
        /// </summary>
        public static Hashtable MResManagers
        {
            get { return _resManagers; }
            set { _resManagers = value; }
        }

        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 ScaleHeight measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleHeight. </returns>
        /// <param name="height">A <see cref="T:System.Double" /> that represents the height.</param>
        /// <param name="scaleHeight">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleHeight for the current ScaleMode.</param>
        /// <param name="originalHeightInPixels">An Integer that represents the height in pixels.</param>
        public static double FromPixelsUserHeight(double height, double scaleHeight, int originalHeightInPixels)
        {
            return height * scaleHeight / originalHeightInPixels;
        }

        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 ScaleWidth measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleWidth. </returns>
        /// <param name="width">A <see cref="T:System.Double" /> that represents the width.</param>
        /// <param name="scaleWidth">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleWidth for the current ScaleMode.</param>
        /// <param name="originalWidthInPixels">An Integer that represents the width in pixels.</param>
        public static double FromPixelsUserWidth(double width, double scaleWidth, int originalWidthInPixels)
        {
            return width * scaleWidth / originalWidthInPixels;
        }


        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 ScaleLeft measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleLeft. </returns>
        /// <param name="x">A <see cref="T:System.Double" /> that represents the X coordinate.</param>
        /// <param name="scaleLeft">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleLeft for the current ScaleMode.</param>
        /// <param name="scaleWidth">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleWidth for the current ScaleMode.</param>
        /// <param name="originalWidthInPixels">An Integer that represents the width in pixels.</param>
        public static double FromPixelsUserX(double x, double scaleLeft, double scaleWidth, int originalWidthInPixels)
        {
            return x * scaleWidth / originalWidthInPixels + scaleLeft;
        }

        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 ScaleTop measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleTop. </returns>
        /// <param name="y">A <see cref="T:System.Double" /> that represents the X coordinate.</param>
        /// <param name="scaleTop">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleTop for the current ScaleMode.</param>
        /// <param name="scaleHeight">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleHeight for the current ScaleMode.</param>
        /// <param name="originalHeightInPixels">An Integer that represents the height in pixels.</param>
        public static double FromPixelsUserY(double y, double scaleTop, double scaleHeight, int originalHeightInPixels)
        {
            return y * scaleHeight / originalHeightInPixels + scaleTop;
        }

        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 measurement for a given ScaleMode.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the Visual Basic 6.0 value for the specified ScaleMode.</returns>
        /// <param name="x">A <see cref="T:System.Double" /> that represents the X coordinate in pixels.</param>
        /// <param name="toScale">A "ScaleMode" enumeration that represents the Visual Basic 6.0 ScaleMode to convert to.</param>
        public static double FromPixelsX(double x, ScaleMode toScale)
        {
            switch (toScale)
            {
                case ScaleMode.Points:
                    return PixelsToTwipsX(x) / 20.0;
                case ScaleMode.Characters:
                    return PixelsToTwipsX(x) / 120.0;
                case ScaleMode.Inches:
                    return PixelsToTwipsX(x) / 1440.0;
                case ScaleMode.Millimeters:
                    return PixelsToTwipsX(x) / 56.692913385826778;
                case ScaleMode.Centimeters:
                    return PixelsToTwipsX(x) / 566.92913385826773;
                case ScaleMode.Himetric:
                    return PixelsToTwipsX(x) / 0.56692913385826771;
            }

            const double result = 0.0;
            return result;
        }

        /// <summary>Converts a pixel measurement to a Visual Basic 6.0 measurement for a given ScaleMode" />.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the Visual Basic 6.0 value for the specified ScaleMode.</returns>
        /// <param name="y">A <see cref="T:System.Double" /> that represents the Y coordinate in pixels.</param>
        /// <param name="toScale">A "ScaleMode" enumeration that represents the Visual Basic 6.0 ScaleMode to convert to.</param>
        public static double FromPixelsY(double y, ScaleMode toScale)
        {
            switch (toScale)
            {
                case ScaleMode.Points:
                    return PixelsToTwipsY(y) / 20.0;
                case ScaleMode.Characters:
                    return PixelsToTwipsY(y) / 240.0;
                case ScaleMode.Inches:
                    return PixelsToTwipsY(y) / 1440.0;
                case ScaleMode.Millimeters:
                    return PixelsToTwipsY(y) / 56.692913385826778;
                case ScaleMode.Centimeters:
                    return PixelsToTwipsY(y) / 566.92913385826773;
                case ScaleMode.Himetric:
                    return PixelsToTwipsY(y) / 0.56692913385826771;
            }

            const double result = 0.0;
            return result;
        }
        private static Image GetImageFromParams(IntPtr handle, int pictype, IntPtr paletteHandle, int width, int height)
        {
            switch (pictype)
            {
                case -1:
                    return null;
                case 0:
                    return null;
                case 1:
                    return Image.FromHbitmap(handle, paletteHandle);
                case 2:
                    Metafile metafile = new Metafile(handle, new WmfPlaceableFileHeader(), false);
                    object objectValue = RuntimeHelpers.GetObjectValue(metafile.Clone());
                    return (Image)objectValue;
                case 4:
                    {
                        Metafile metafile2 = new Metafile(handle, false);
                        object objectValue2 = RuntimeHelpers.GetObjectValue(metafile2.Clone());
                        return (Image)objectValue2;
                    }
            }
            throw new NotSupportedException();
        }

        /// <summary>Gets a  for a given OLE IPicture object.</summary>
        /// <returns>A <see cref="T:System.Drawing.Image" />.</returns>
        /// <param name="pict">An OLE IPicture object.</param>
        // ReSharper disable InconsistentNaming
        public static Image IPictureToImage(object pict)
        // ReSharper restore InconsistentNaming
        {
            if (pict == null)
            {
                return null;
            }
            IntPtr zero = IntPtr.Zero;
            // ReSharper disable SuggestUseVarKeywordEvident
            IPicture picture = (IPicture)pict;
            // ReSharper restore SuggestUseVarKeywordEvident
            int type = picture.Type;
            if (type == 1)
            {
                try
                {
                    zero = new IntPtr(picture.hPal);
                }
                catch (COMException expr_2C)
                {
                    ProjectData.SetProjectError(expr_2C);
                    ProjectData.ClearProjectError();
                }
            }

            IntPtr handle = new IntPtr(picture.Handle);
            return GetImageFromParams(handle, type, zero, picture.Width, picture.Height);
        }

        /// <summary>Converts an X coordinate from pixels to twips.</summary>
        /// <returns>A Double that contains the X coordinate expressed in twips.</returns>
        /// <param name="x">A Double that contains the X coordinate to convert.</param>
        public static double PixelsToTwipsX(double x)
        {
            SetUpTwipsPerPixel();
            return x * _twipsPerPixelX;
        }

        /// <summary>Converts a Y coordinate from pixels to twips.</summary>
        /// <returns>A Double that contains the Y coordinate expressed in twips.</returns>
        /// <param name="y">A Double that contains the Y coordinate to convert.</param>
        public static double PixelsToTwipsY(double y)
        {
            SetUpTwipsPerPixel();
            return y * _twipsPerPixelY;
        }


        private static void SetUpTwipsPerPixel()
        {
            if (!_isTwipsPerPixelSetUp)
            {
                _twipsPerPixelX = 0.0;
                _twipsPerPixelY = 0.0;
                try
                {
                    IntPtr dC = SupportHelperNativeMethods.GetDC(SupportHelperNativeMethods.NullIntPtr);
                    if (!dC.Equals(SupportHelperNativeMethods.NullIntPtr))
                    {
                        _twipsPerPixelX = 1440.0 / SupportHelperNativeMethods.GetDeviceCaps(dC, 88);
                        _twipsPerPixelY = 1440.0 / SupportHelperNativeMethods.GetDeviceCaps(dC, 90);
                        SupportHelperNativeMethods.ReleaseDC(SupportHelperNativeMethods.NullIntPtr, dC);
                    }
                }
                catch (Exception arg_8B0)
                {
                    ProjectData.SetProjectError(arg_8B0);
                    ProjectData.ClearProjectError();
                }
                _isTwipsPerPixelSetUp = true;
                // ReSharper disable CompareOfFloatsByEqualityOperator
                if (_twipsPerPixelX == 0.0 || _twipsPerPixelY == 0.0)
                // ReSharper restore CompareOfFloatsByEqualityOperator
                {
                    _twipsPerPixelX = 15.0;
                    _twipsPerPixelY = 15.0;
                }
            }
        }

        /// <summary>Displays a form by calling either the <see cref="M:System.Windows.Forms.Control.Show" /> or <see cref="M:System.Windows.Forms.Form.ShowDialog" /> method.</summary>
        /// <param name="form">The <see cref="T:System.Windows.Forms.Form" /> to display.</param>
        /// <param name="modal">Optional. A enumeration that specifies modality.</param>
        /// <param name="ownerForm">Optional. The parameter of the method.</param>
        public static void ShowForm(Form form, int modal, Form ownerForm)
        {
            if (ownerForm != null)
            {
                form.Owner = ownerForm;
            }
            if (modal == 0)
            {
                form.Show();
            }
            else
            {
                if (modal == 1)
                {
                    form.ShowDialog();
                }
                else
                {
                    throw new NotSupportedException();
                }
            }
        }

        /// <summary>Converts a Visual Basic 6.0 ScaleHeight measurement to a pixel measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleHeight. </returns>
        /// <param name="height">A <see cref="T:System.Double" /> that represents the height.</param>
        /// <param name="scaleHeight">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleHeight for the current ScaleMode.</param>
        /// <param name="originalHeightInPixels">An Integer that represents the height in pixels.</param>
        public static double ToPixelsUserHeight(double height, double scaleHeight, int originalHeightInPixels)
        {
            return height / scaleHeight * originalHeightInPixels;
        }

        /// <summary>Converts a Visual Basic 6.0 ScaleWidth measurement to a pixel measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleWidth. </returns>
        /// <param name="width">A <see cref="T:System.Double" /> that represents the width.</param>
        /// <param name="scaleWidth">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleWidth for the current ScaleMode.</param>
        /// <param name="originalWidthInPixels">An Integer that represents the width in pixels.</param>
        public static double ToPixelsUserWidth(double width, double scaleWidth, int originalWidthInPixels)
        {
            return width / scaleWidth * originalWidthInPixels;
        }

        /// <summary>Converts a Visual Basic 6.0 ScaleLeft measurement to a pixel measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleLeft. </returns>
        /// <param name="x">A <see cref="T:System.Double" /> that represents the X coordinate.</param>
        /// <param name="scaleLeft">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleLeft for the current ScaleMode.</param>
        /// <param name="scaleWidth">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleWidth for the current ScaleMode.</param>
        /// <param name="originalWidthInPixels">An Integer that represents the width in pixels.</param>
        public static double ToPixelsUserX(double x, double scaleLeft, double scaleWidth, int originalWidthInPixels)
        {
            return (x - scaleLeft) / scaleWidth * originalWidthInPixels;
        }

        /// <summary>Converts a Visual Basic 6.0 ScaleTop measurement to a pixel measurement.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the converted Visual Basic 6.0 ScaleLeft. </returns>
        /// <param name="y">A <see cref="T:System.Double" /> that represents the Y coordinate.</param>
        /// <param name="scaleTop">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleTop for the current ScaleMode.</param>
        /// <param name="scaleHeight">A <see cref="T:System.Double" /> that represents the Visual Basic 6.0 ScaleHeight for the current ScaleMode.</param>
        /// <param name="originalHeightInPixels">An Integer that represents the height in pixels.</param>
        public static double ToPixelsUserY(double y, double scaleTop, double scaleHeight, int originalHeightInPixels)
        {
            return (y - scaleTop) / scaleHeight * originalHeightInPixels;
        }

        /// <summary>Converts a Visual Basic 6.0 measurement to a pixel measurement for a given.</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the pixel value for the specified ScaleMode.</returns>
        /// <param name="x">A <see cref="T:System.Double" /> that represents the X coordinate.</param>
        /// <param name="fromScale">A enumeration that represents the Visual Basic 6.0 ScaleMode to convert from.</param>
        public static double ToPixelsX(double x, ScaleMode fromScale)
        {
            switch (fromScale)
            {
                case ScaleMode.Points:
                    return TwipsToPixelsX(x * 20.0);
                case ScaleMode.Characters:
                    return TwipsToPixelsX(x * 120.0);
                case ScaleMode.Inches:
                    return TwipsToPixelsX(x * 1440.0);
                case ScaleMode.Millimeters:
                    return TwipsToPixelsX(x * 56.692913385826778);
                case ScaleMode.Centimeters:
                    return TwipsToPixelsX(x * 566.92913385826773);
                case ScaleMode.Himetric:
                    return TwipsToPixelsX(x * 0.56692913385826771);
            }

            const double result = 0.0;
            return result;
        }

        /// <summary>Converts a Visual Basic 6.0 measurement to a pixel measurement for a given .</summary>
        /// <returns>A <see cref="T:System.Double" /> that contains the pixel value for the specified ScaleMode.</returns>
        /// <param name="y">A <see cref="T:System.Double" /> that represents the Y coordinate.</param>
        /// <param name="fromScale">A enumeration that represents the Visual Basic 6.0 ScaleMode to convert from.</param>
        public static double ToPixelsY(double y, ScaleMode fromScale)
        {
            switch (fromScale)
            {
                case ScaleMode.Points:
                    return TwipsToPixelsY(y * 20.0);
                case ScaleMode.Characters:
                    return TwipsToPixelsY(y * 240.0);
                case ScaleMode.Inches:
                    return TwipsToPixelsY(y * 1440.0);
                case ScaleMode.Millimeters:
                    return TwipsToPixelsY(y * 56.692913385826778);
                case ScaleMode.Centimeters:
                    return TwipsToPixelsY(y * 566.92913385826773);
                case ScaleMode.Himetric:
                    return TwipsToPixelsY(y * 0.56692913385826771);
            }


            const double result = 0.0;
            return result;
        }

        /// <summary>Gets a value that is used to convert twips to pixels based on screen settings.</summary>
        /// <returns>A Double that contains the conversion factor.</returns>
        public static float TwipsPerPixelX()
        {
            SetUpTwipsPerPixel();
            return (float)_twipsPerPixelX;
        }

        /// <summary>Gets a value that is used to convert twips to pixels based on screen settings.</summary>
        /// <returns>A Double that contains the conversion factor.</returns>
        public static float TwipsPerPixelY()
        {
            SetUpTwipsPerPixel();
            return (float)_twipsPerPixelY;
        }

        /// <summary>Converts an X coordinate from twips to pixels.</summary>
        /// <returns>A Double that contains the X coordinate expressed in pixels.</returns>
        /// <param name="x">A Double that contains the X coordinate to convert.</param>
        public static double TwipsToPixelsX(double x)
        {
            SetUpTwipsPerPixel();
            return x / _twipsPerPixelX;
        }

        /// <summary>Converts a Y coordinate from twips to pixels.</summary>
        /// <returns>A Double that contains the Y coordinate expressed in pixels.</returns>
        /// <param name="y">A Double that contains the X coordinate to convert.</param>
        public static double TwipsToPixelsY(double y)
        {
            SetUpTwipsPerPixel();
            return y / _twipsPerPixelY;
        }


        internal static class SupportHelperNativeMethods
        {
            internal static IntPtr NullIntPtr = new IntPtr(0);

            [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true)]
            internal static extern int ReleaseDC(IntPtr hWnd, IntPtr hDc);

            [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true)]
            internal static extern IntPtr GetDC(IntPtr hWnd);

            [DllImport("gdi32", CharSet = CharSet.Auto, SetLastError = true)]
            internal static extern int GetDeviceCaps(IntPtr hDc, int nIndex);
        }
    }


}

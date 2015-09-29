namespace RP.Util.Icon
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Windows;
    using System.Windows.Interop;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;

    public static class IconExtensions
    {
        #region Public Methods

        public static Icon GetAssociatedIconLarge(this Process process)
        {
            string path = process.StartInfo.FileName;
            if (string.IsNullOrEmpty(path))
            {
                try
                {
                    path = process.MainModule.FileName;
                }
                catch { }
            }

            Icon icon = null;
            if (!string.IsNullOrEmpty(path))
            {
                icon = IconExtractor.GetAssociatedIconLarge(process.StartInfo.FileName);
            }
            return icon;
        }

        public static Icon GetAssociatedIconSmall(this Process process)
        {
            string path = process.StartInfo.FileName;
            if (string.IsNullOrEmpty(path))
            {
                try
                {
                    path = process.MainModule.FileName;
                }
                catch { }
            }

            Icon icon = null;
            if (!string.IsNullOrEmpty(path))
            {
                icon = IconExtractor.GetAssociatedIconSmall(process.StartInfo.FileName);
            }
            return icon;
        }

        public static ImageSource ToImageSource(this Icon icon)
        {
            ImageSource wpfBitmap = null;

            if (icon != null)
            {
                Bitmap bitmap = icon.ToBitmap();
                IntPtr hBitmap = bitmap.GetHbitmap();

                wpfBitmap = Imaging.CreateBitmapSourceFromHBitmap
                  (
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

                if (!IconExtractor.DeleteObject(hBitmap))
                {
                    throw new Win32Exception();
                }
            }

            return wpfBitmap;
        }

        #endregion
    }
}
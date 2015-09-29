namespace RP.Util.Icon
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;

    public static class IconExtractor
    {
        #region Constants

        private const uint iconFlag = 0x100;
        private const uint largeIconFlag = 0x0; // 'Large icon
        private const uint smallIconFlag = 0x1; // 'Small icon

        #endregion

        #region Public Methods

        public static int CountIcons(string path)
        {
            return ExtractIconEx(path, -1, null, null, 0);
        }

        public static System.Drawing.Icon GetAssociatedIconLarge(string path)
        {
            FileInfo info = new FileInfo();

            SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf(info), iconFlag | largeIconFlag);

            System.Drawing.Icon toReturn = null;

            if (info.Icon != IntPtr.Zero)
            {
                toReturn = System.Drawing.Icon.FromHandle(info.Icon);
            }

            return toReturn;
        }

        public static System.Drawing.Icon GetAssociatedIconSmall(string path)
        {
            FileInfo info = new FileInfo();

            SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf(info), iconFlag | smallIconFlag);

            System.Drawing.Icon toReturn = null;

            if (info.Icon != IntPtr.Zero)
            {
                toReturn = System.Drawing.Icon.FromHandle(info.Icon);
            }

            return toReturn;
        }

        public static System.Drawing.Icon[] GetLargeIcons(string path)
        {
            IntPtr[] icons;
            IntPtr[] dummy;
            GetIcons(path, out icons, out dummy);

            List<System.Drawing.Icon> iconList = new List<System.Drawing.Icon>();

            foreach (IntPtr ptr in icons)
            {
                System.Drawing.Icon icon = System.Drawing.Icon.FromHandle(ptr);
                iconList.Add((System.Drawing.Icon)icon.Clone());
                DestroyIcon(ptr);
            }

            return iconList.ToArray();
        }

        public static System.Drawing.Icon[] GetSmallIcons(string path)
        {
            IntPtr[] icons;
            IntPtr[] dummy;
            GetIcons(path, out dummy, out icons);

            List<System.Drawing.Icon> iconList = new List<System.Drawing.Icon>();

            foreach (IntPtr ptr in icons)
            {
                System.Drawing.Icon icon = System.Drawing.Icon.FromHandle(ptr);
                iconList.Add((System.Drawing.Icon)icon.Clone());
                DestroyIcon(ptr);
            }

            return iconList.ToArray();
        }

        #endregion

        #region Protected/Internal Methods

        [DllImport("gdi32.dll", SetLastError = true)]
        internal static extern bool DeleteObject(IntPtr hObject);

        #endregion

        #region Private Methods

        [DllImport("user32.dll", EntryPoint = "DestroyIcon", SetLastError = true)]
        private static extern int DestroyIcon(IntPtr icon);

        [DllImport("Shell32", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string path, int iconIndex, IntPtr[] largeIcons, IntPtr[] smallIcons, int iconCount);

        private static void GetIcons(string path, out IntPtr[] large, out IntPtr[] small)
        {
            int count = CountIcons(path);
            large = new IntPtr[count];
            small = new IntPtr[count];

            ExtractIconEx(path, 0, large, small, count);
        }

        [DllImport("shell32.dll")]
        private static extern IntPtr SHGetFileInfo(string path, uint fileAttributes, ref FileInfo info, uint size, uint flags);

        #endregion

        #region Nested type: FileInfo

        [StructLayout(LayoutKind.Sequential)]
        private struct FileInfo
        {
            public readonly IntPtr Icon;
            public readonly IntPtr Index;
            public readonly uint Attributes;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public readonly string DisplayName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public readonly string TypeName;
        };

        #endregion
    }
}
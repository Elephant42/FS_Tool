using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;

namespace HidSharp
{
	/*
	 * Please leave this Copyright notice in your code if you use it
	 * Written by Decebal Mihailescu [http://www.codeproject.com/script/articles/list_articles.asp?userid=634640]
	 */

[StructLayout(LayoutKind.Sequential)]
	public struct RECT
	{
		public int Left;
		public int Top;
		public int Right;
		public int Bottom;
		public override string ToString()
		{
			return string.Format("Left = {0}, Top = {1}, Right = {2}, Bottom ={3}",
				Left, Top, Right, Bottom);
		}
		public int Width
		{
			get { return Math.Abs(Right - Left); }
		}
		public int Height
		{
			get { return Math.Abs(Bottom - Top); }
		}
	}

	public static class WindowUtil
    {

		public static void SetWindowFocus(String windowText)
        {
			IntPtr hWin = Win32API.FindWindow(null, windowText);

			SetWindowFocus(hWin);
		}
		public static void SetWindowFocus(IntPtr windowHandle)
		{
			Win32API.SetForegroundWindow(windowHandle);
		}

		public static Bitmap GetWindowSnapshot(IntPtr AppWndHandle, bool IsClientWnd, Win32API.WindowShowStyle nCmdShow)
		{
			if (AppWndHandle == IntPtr.Zero || !Win32API.IsWindow(AppWndHandle) ||
						!Win32API.IsWindowVisible(AppWndHandle))
				return null;
			if (Win32API.IsIconic(AppWndHandle))
				Win32API.ShowWindow(AppWndHandle, nCmdShow);//show it
			if (!Win32API.SetForegroundWindow(AppWndHandle))
				return null;//can't bring it to front
			System.Threading.Thread.Sleep(1000);//give it some time to redraw
			RECT appRect;
			bool res = IsClientWnd ? Win32API.GetClientRect
				(AppWndHandle, out appRect) : Win32API.GetWindowRect
				(AppWndHandle, out appRect);
			if (!res || appRect.Height == 0 || appRect.Width == 0)
			{
				return null;//some hidden window
			}
			// calculate the app rectangle
			if (IsClientWnd)
			{
				Point lt = new Point(appRect.Left, appRect.Top);
				Point rb = new Point(appRect.Right, appRect.Bottom);
				Win32API.ClientToScreen(AppWndHandle, ref lt);
				Win32API.ClientToScreen(AppWndHandle, ref rb);
				appRect.Left = lt.X;
				appRect.Top = lt.Y;
				appRect.Right = rb.X;
				appRect.Bottom = rb.Y;
			}
			//Intersect with the Desktop rectangle and get what's visible
			IntPtr DesktopHandle = Win32API.GetDesktopWindow();
			RECT desktopRect;
			Win32API.GetWindowRect(DesktopHandle, out desktopRect);
			RECT visibleRect;
			if (!Win32API.IntersectRect
				(out visibleRect, ref desktopRect, ref appRect))
			{
				visibleRect = appRect;
			}
			if (Win32API.IsRectEmpty(ref visibleRect))
				return null;

			int Width = visibleRect.Width;
			int Height = visibleRect.Height;
			IntPtr hdcTo = IntPtr.Zero;
			IntPtr hdcFrom = IntPtr.Zero;
			IntPtr hBitmap = IntPtr.Zero;
			try
			{
				Bitmap clsRet = null;

				// get device context of the window...
				hdcFrom = IsClientWnd ? Win32API.GetDC(AppWndHandle) :
						Win32API.GetWindowDC(AppWndHandle);

				// create dc that we can draw to...
				hdcTo = Win32API.CreateCompatibleDC(hdcFrom);
				hBitmap = Win32API.CreateCompatibleBitmap(hdcFrom, Width, Height);

				//  validate
				if (hBitmap != IntPtr.Zero)
				{
					// adjust and copy
					int x = appRect.Left < 0 ? -appRect.Left : 0;
					int y = appRect.Top < 0 ? -appRect.Top : 0;
					IntPtr hLocalBitmap = Win32API.SelectObject(hdcTo, hBitmap);
					Win32API.BitBlt(hdcTo, 0, 0, Width, Height,
						hdcFrom, x, y, Win32API.SRCCOPY);
					Win32API.SelectObject(hdcTo, hLocalBitmap);
					//  create bitmap for window image...
					clsRet = System.Drawing.Image.FromHbitmap(hBitmap);
				}
				return clsRet;
			}
			finally
			{
				//  release the unmanaged resources
				if (hdcFrom != IntPtr.Zero)
					Win32API.ReleaseDC(AppWndHandle, hdcFrom);
				if (hdcTo != IntPtr.Zero)
					Win32API.DeleteDC(hdcTo);
				if (hBitmap != IntPtr.Zero)
					Win32API.DeleteObject(hBitmap);
			}
		}
	}

public	static class Win32API
	{
		[DllImport("User32.Dll")]
		public static extern void GetClassName(IntPtr hWnd, System.Text.StringBuilder param, int length);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern int GetWindowTextLength(IntPtr hWnd);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

		public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

		[DllImport("user32.dll")]
		public static extern bool EnumThreadWindows(uint dwThreadId, Win32API.EnumThreadDelegate lpfn, IntPtr lParam);

		[DllImport("user32.dll")]
		public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

		[DllImport("user32.dll")]
		public static extern bool AppendMenu(IntPtr hMenu, Int32 wFlags, Int32 wIDNewItem, string lpNewItem);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool ShowWindow(IntPtr hWnd, WindowShowStyle nCmdShow);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsRectEmpty([In] ref RECT lprc);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool ClientToScreen(IntPtr hwnd, ref Point lpPoint);

		public const Int32 WM_SYSCOMMAND = 0x112;
		public const Int32 MF_SEPARATOR = 0x800;
		public const Int32 MF_STRING = 0x0;

		/// <summary>Enumeration of the different ways of showing a window using 
		/// ShowWindow</summary>
		public enum WindowShowStyle : uint
		{
			/// <summary>Hides the window and activates another window.</summary>
			/// <remarks>See SW_HIDE</remarks>
			Hide = 0,
			/// <summary>Activates and displays a window. If the window is minimized 
			/// or maximized, the system restores it to its original size and 
			/// position. An application should specify this flag when displaying 
			/// the window for the first time.</summary>
			/// <remarks>See SW_SHOWNORMAL</remarks>
			ShowNormal = 1,
			/// <summary>Activates the window and displays it as a minimized window.</summary>
			/// <remarks>See SW_SHOWMINIMIZED</remarks>
			ShowMinimized = 2,
			/// <summary>Activates the window and displays it as a maximized window.</summary>
			/// <remarks>See SW_SHOWMAXIMIZED</remarks>
			ShowMaximized = 3,
			/// <summary>Maximizes the specified window.</summary>
			/// <remarks>See SW_MAXIMIZE</remarks>
			Maximize = 3,
			/// <summary>Displays a window in its most recent size and position. 
			/// This value is similar to "ShowNormal", except the window is not 
			/// actived.</summary>
			/// <remarks>See SW_SHOWNOACTIVATE</remarks>
			ShowNormalNoActivate = 4,
			/// <summary>Activates the window and displays it in its current size 
			/// and position.</summary>
			/// <remarks>See SW_SHOW</remarks>
			Show = 5,
			/// <summary>Minimizes the specified window and activates the next 
			/// top-level window in the Z order.</summary>
			/// <remarks>See SW_MINIMIZE</remarks>
			Minimize = 6,
			/// <summary>Displays the window as a minimized window. This value is 
			/// similar to "ShowMinimized", except the window is not activated.</summary>
			/// <remarks>See SW_SHOWMINNOACTIVE</remarks>
			ShowMinNoActivate = 7,
			/// <summary>Displays the window in its current size and position. This 
			/// value is similar to "Show", except the window is not activated.</summary>
			/// <remarks>See SW_SHOWNA</remarks>
			ShowNoActivate = 8,
			/// <summary>Activates and displays the window. If the window is 
			/// minimized or maximized, the system restores it to its original size 
			/// and position. An application should specify this flag when restoring 
			/// a minimized window.</summary>
			/// <remarks>See SW_RESTORE</remarks>
			Restore = 9,
			/// <summary>Sets the show state based on the SW_ value specified in the 
			/// STARTUPINFO structure passed to the CreateProcess function by the 
			/// program that started the application.</summary>
			/// <remarks>See SW_SHOWDEFAULT</remarks>
			ShowDefault = 10,
			/// <summary>Windows 2000/XP: Minimizes a window, even if the thread 
			/// that owns the window is hung. This flag should only be used when 
			/// minimizing windows from a different thread.</summary>
			/// <remarks>See SW_FORCEMINIMIZE</remarks>
			ForceMinimized = 11
		}

		public const int SRCCOPY = 13369376;

		[DllImport("user32.dll", SetLastError = true)]
		public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", EntryPoint = "GetDC")]
		public extern static IntPtr GetDC(IntPtr hWnd);

		[DllImport("user32.dll", EntryPoint = "ReleaseDC")]
		public extern static IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);

		[DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
		public extern static IntPtr DeleteDC(IntPtr hDc);

		[DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
		public extern static IntPtr DeleteObject(IntPtr hDc);

		[DllImport("gdi32.dll", EntryPoint = "BitBlt")]
		public extern static bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, int RasterOp);

		[DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
		public extern static IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

		[DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC")]
		public extern static IntPtr CreateCompatibleDC(IntPtr hdc);

		[DllImport("gdi32.dll", EntryPoint = "SelectObject")]
		public extern static IntPtr SelectObject(IntPtr hdc, IntPtr bmp);

		[DllImport("user32.dll", SetLastError = false)]
		public static extern IntPtr GetDesktopWindow();

		[DllImport("user32.dll")]
		public static extern IntPtr GetWindowDC(IntPtr hWnd);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool SetForegroundWindow(IntPtr hWnd);

		[DllImport("User32.Dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsIconic(IntPtr hWnd);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindowVisible(IntPtr hWnd);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

		[DllImport("user32.dll", SetLastError = true)]
		public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IntersectRect(out RECT lprcDst, [In] ref RECT lprcSrc1,
		   [In] ref RECT lprcSrc2);
	}
}

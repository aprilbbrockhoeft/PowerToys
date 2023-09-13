// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using ManagedCommon;
using Microsoft.PowerToys.Telemetry;

namespace PowerOCR.Utilities;

public static class WindowUtilities
{
    public static void LaunchOCROverlayOnEveryScreen()
    {
        if (IsOCROverlayCreated())
        {
            Logger.LogWarning("Tried to launch the overlay, but it has been already created.");
            return;
        }

        Logger.LogInfo($"Adding Overlays for each screen");
        IntPtr foregroundWindowHandle = GetForegroundWindow();
        GetClientRect(foregroundWindowHandle, out RECT foregroundWindowRect);
        var foregroundWindowRectangle = RECTToRectangle(foregroundWindowRect);
        var screenPoint = new POINT(foregroundWindowRectangle.Left, foregroundWindowRectangle.Top);
        ClientToScreen(foregroundWindowHandle, ref screenPoint);
        foregroundWindowRectangle.Location = screenPoint;
        var sb = new StringBuilder(1000);
        var res = GetWindowText(foregroundWindowHandle, sb, 1000);
        Logger.LogInfo($"{foregroundWindowRectangle.Left}, {foregroundWindowRectangle.Top}, {foregroundWindowRectangle.Right}, {foregroundWindowRectangle.Bottom}, {foregroundWindowRectangle.Width}, {foregroundWindowRectangle.Height}");
        Logger.LogInfo($"foreground window {sb.ToString()}");
        foreach (Screen screen in Screen.AllScreens)
        {
            Logger.LogInfo($"screen {screen}");
            OCROverlay overlay = new(screen.Bounds, foregroundWindowRectangle);

            overlay.Show();
            ActivateWindow(overlay);
        }

        PowerToysTelemetry.Log.WriteEvent(new PowerOCR.Telemetry.PowerOCRInvokedEvent());
    }

    internal static bool IsOCROverlayCreated()
    {
        WindowCollection allWindows = System.Windows.Application.Current.Windows;

        foreach (Window window in allWindows)
        {
            if (window is OCROverlay)
            {
                return true;
            }
        }

        return false;
    }

    internal static void CloseAllOCROverlays()
    {
        WindowCollection allWindows = System.Windows.Application.Current.Windows;

        foreach (Window window in allWindows)
        {
            if (window is OCROverlay overlay)
            {
                overlay.Close();
            }
        }

        GC.Collect();

        // TODO: Decide when to close the process
        // System.Windows.Application.Current.Shutdown();
    }

    public static void ActivateWindow(Window window)
    {
        var handle = new System.Windows.Interop.WindowInteropHelper(window).Handle;
        var fgHandle = OSInterop.GetForegroundWindow();

        var threadId1 = OSInterop.GetWindowThreadProcessId(handle, System.IntPtr.Zero);
        var threadId2 = OSInterop.GetWindowThreadProcessId(fgHandle, System.IntPtr.Zero);

        if (threadId1 != threadId2)
        {
            OSInterop.AttachThreadInput(threadId1, threadId2, true);
            OSInterop.SetForegroundWindow(handle);
            OSInterop.AttachThreadInput(threadId1, threadId2, false);
        }
        else
        {
            OSInterop.SetForegroundWindow(handle);
        }
    }

    internal static void OcrOverlayKeyDown(Key key, bool? isActive = null)
    {
        WindowCollection allWindows = System.Windows.Application.Current.Windows;

        if (key == Key.Escape)
        {
            PowerToysTelemetry.Log.WriteEvent(new PowerOCR.Telemetry.PowerOCRCancelledEvent());
            CloseAllOCROverlays();
        }

        foreach (Window window in allWindows)
        {
            if (window is OCROverlay overlay)
            {
                overlay.KeyPressed(key, isActive);
            }
        }
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    private static Rectangle RECTToRectangle(RECT r)
    {
        return new Rectangle(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }

        public static implicit operator System.Drawing.Point(POINT p)
        {
            return new System.Drawing.Point(p.X, p.Y);
        }

        public static implicit operator POINT(System.Drawing.Point p)
        {
            return new POINT(p.X, p.Y);
        }

        public override string ToString()
        {
            return $"X: {X}, Y: {Y}";
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
    }
}

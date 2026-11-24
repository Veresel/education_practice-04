using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using DrawingPoint = System.Drawing.Point;
using FormsSendKeys = System.Windows.Forms.SendKeys;
using IOPath = System.IO.Path;

internal static class Program
{
    private static string phase = "before";
    private static string shopDir = "";
    private static string evidenceDir = "";
    private static Process? app;

    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }

    public static int Main(string[] args)
    {
        phase = args.Length > 0 ? args[0] : "before";
        shopDir = IOPath.GetFullPath(args.Length > 1 ? args[1] : @"..\..\..\..\stands\Shop-source");
        evidenceDir = IOPath.GetFullPath(IOPath.Combine(shopDir, "..", "..", "evidence"));
        Directory.CreateDirectory(evidenceDir);

        RestoreWork();
        StartApp();
        try
        {
            var win = WaitMain();
            if (phase == "before") CaptureBefore(win);
            else CaptureAfter(win);
        }
        finally
        {
            StopApp();
            RestoreWork();
        }
        return 0;
    }

    private static void CaptureBefore(AutomationElement win)
    {
        Shot(win, "01-control-state");
        SetValue(FindId(win, "LoginBox"), "nobody");
        Invoke(FindId(win, "LoginButton"));
        Thread.Sleep(400);
        ShotDesktop("02-d1-nobody-before");
        DismissDialogs();

        RestoreWork();
        Restart(ref win);
        Login(win, "client");
        DismissDialogs();
        OpenProduct(win, "Электрод pH");
        SetValue(FindId(win, "QuantityBox"), "1");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        Thread.Sleep(300);
        Shot(win, "05-d2-cart-before-back");
        Invoke(FindId(win, "BackButton"));
        Thread.Sleep(400);
        OpenProduct(win, "Электрод pH — резервный узел");
        Thread.Sleep(300);
        Shot(win, "06-d2-cart-after-back-before-fix");

        RestoreWork();
        Restart(ref win);
        Login(win, "client");
        DismissDialogs();
        ApplyFilter(win, "сервисный", "Электроды", 0);
        Thread.Sleep(400);
        Shot(win, "08-d3-filter-before");
    }

    private static void CaptureAfter(AutomationElement win)
    {
        Shot(win, "01-control-state");
        SetValue(FindId(win, "LoginBox"), "nobody");
        Invoke(FindId(win, "LoginButton"));
        Thread.Sleep(400);
        ShotDialogOrMain(win, "03-d1-nobody-after");
        DismissDialogs();
        SetValue(FindId(win, "LoginBox"), "client");
        Invoke(FindId(win, "LoginButton"));
        Thread.Sleep(400);
        DismissDialogs();
        Shot(win, "04-d1-client-after");

        OpenProduct(win, "Электрод pH");
        SetValue(FindId(win, "QuantityBox"), "1");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        Shot(win, "05-d2-cart-before-back");
        Invoke(FindId(win, "BackButton"));
        Thread.Sleep(300);
        OpenProduct(win, "Электрод pH — резервный узел");
        Thread.Sleep(300);
        Shot(win, "07-d2-cart-after-back-after-fix");

        Invoke(FindId(win, "BackButton"));
        Thread.Sleep(300);
        ApplyFilter(win, "сервисный", "Электроды", 0);
        Thread.Sleep(400);
        Shot(win, "09-d3-filter-after");

        RestoreWork();
        Restart(ref win);
        Login(win, "client");
        DismissDialogs();
        OpenProduct(win, "Электрод pH");
        SetValue(FindId(win, "QuantityBox"), "3");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        SetValue(FindId(win, "QuantityBox"), "3");
        Invoke(FindId(win, "AddButton"));
        Thread.Sleep(400);
        ShotDialogOrMain(win, "10-quantity-reject");
        DismissDialogs();
        SetValue(FindId(win, "QuantityBox"), "2");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        Thread.Sleep(300);
        Shot(win, "11-quantity-continue");

        RestoreWork();
        Restart(ref win);
        Login(win, "client");
        DismissDialogs();
        OpenProduct(win, "Электрод pH");
        SetValue(FindId(win, "QuantityBox"), "2");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        Invoke(FindId(win, "BackButton"));
        Thread.Sleep(300);
        OpenProduct(win, "Преобразователь сигнала");
        SetValue(FindId(win, "QuantityBox"), "1");
        Invoke(FindId(win, "AddButton"));
        DismissDialogs();
        Thread.Sleep(300);
        Shot(win, "12-two-lines-before-confirm");
        Invoke(FindId(win, "ConfirmButton"));
        Thread.Sleep(500);
        ShotDialogOrMain(win, "13-confirm-result");
        DismissDialogs();
        StopApp();
        StartApp();
        win = WaitMain();
        Shot(win, "14-after-restart");

        RestoreWork();
        Restart(ref win);
        Login(win, "client");
        DismissDialogs();
        ApplyFilter(win, "Электрод", "Все категории", 0);
        Thread.Sleep(400);
        Shot(win, "15-extra-partial-name");
        ApplyFilter(win, "Электрод pH", "Все категории", 0);
        Thread.Sleep(400);
        Shot(win, "15b-extra-full-name");
    }

    private static void RestoreWork()
    {
        File.Copy(IOPath.Combine(shopDir, "baseline.json"), IOPath.Combine(shopDir, "work-data.json"), true);
    }

    private static void StartApp()
    {
        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = "run --project Shop.csproj -- work-data.json",
            WorkingDirectory = shopDir,
            UseShellExecute = false
        };
        app = Process.Start(psi) ?? throw new InvalidOperationException("Не удалось запустить приложение.");
    }

    private static void Restart(ref AutomationElement win)
    {
        StopApp();
        StartApp();
        win = WaitMain();
    }

    private static void StopApp()
    {
        try
        {
            if (app is { HasExited: false })
            {
                app.Kill(true);
                app.WaitForExit(5000);
            }
        }
        catch { }
        foreach (var p in Process.GetProcessesByName("Shop"))
        {
            try { p.Kill(true); } catch { }
        }
        Thread.Sleep(500);
        app = null;
    }

    private static AutomationElement WaitMain()
    {
        for (var i = 0; i < 60; i++)
        {
            foreach (var p in Process.GetProcesses())
            {
                try
                {
                    if (p.ProcessName.Contains("Shop", StringComparison.OrdinalIgnoreCase) && p.MainWindowHandle != IntPtr.Zero)
                    {
                        ShowWindow(p.MainWindowHandle, 3);
                        SetForegroundWindow(p.MainWindowHandle);
                        var el = AutomationElement.FromHandle(p.MainWindowHandle);
                        if (el != null) return el;
                    }
                }
                catch { }
            }
            Thread.Sleep(500);
        }
        throw new TimeoutException("Окно приложения не появилось.");
    }

    private static void Login(AutomationElement win, string login)
    {
        SetValue(FindId(win, "LoginBox"), login);
        Invoke(FindId(win, "LoginButton"));
        Thread.Sleep(300);
    }

    private static void ApplyFilter(AutomationElement win, string query, string category, int sortIndex)
    {
        SetValue(FindId(win, "SearchBox"), query);
        SelectCombo(FindId(win, "CategoryBox"), category);
        SelectComboIndex(FindId(win, "SortBox"), sortIndex);
    }

    private static void OpenProduct(AutomationElement win, string exactName)
    {
        var grid = FindId(win, "ProductsGrid");
        var rows = grid.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "DataGridRow"));
        foreach (AutomationElement row in rows)
        {
            var texts = row.FindAll(TreeScope.Children, Condition.TrueCondition);
            var names = new List<string>();
            foreach (AutomationElement t in row.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)))
                names.Add(t.Current.Name);
            if (names.Count == 0) continue;
            var title = names[0];
            var wantReserve = exactName.Contains("резервный", StringComparison.Ordinal);
            var isReserve = title.Contains("резервный", StringComparison.Ordinal);
            if (title != exactName && !(title.Contains(exactName, StringComparison.Ordinal) && isReserve == wantReserve))
                continue;
            if (isReserve != wantReserve) continue;
            var btn = row.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Открыть позицию"));
            if (btn == null) continue;
            Invoke(btn);
            Thread.Sleep(400);
            return;
        }
        throw new InvalidOperationException("Не найдена позиция: " + exactName);
    }

    private static IEnumerable<AutomationElement> Ancestors(AutomationElement el)
    {
        var cur = el;
        while (cur != null)
        {
            yield return cur;
            cur = TreeWalker.ControlViewWalker.GetParent(cur);
        }
    }

    private static IEnumerable<AutomationElement> Descendants(AutomationElement el)
    {
        var all = el.FindAll(TreeScope.Descendants, Condition.TrueCondition);
        foreach (AutomationElement x in all) yield return x;
    }

    private static AutomationElement FindId(AutomationElement root, string id)
    {
        for (var i = 0; i < 20; i++)
        {
            var el = root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (el != null) return el;
            Thread.Sleep(150);
        }
        throw new InvalidOperationException("Не найден элемент " + id);
    }

    private static void SetValue(AutomationElement el, string value)
    {
        el.SetFocus();
        Thread.Sleep(80);
        if (el.TryGetCurrentPattern(ValuePattern.Pattern, out var p) && p is ValuePattern vp)
        {
            vp.SetValue(value);
            return;
        }
        FormsSendKeys.SendWait("^a");
        FormsSendKeys.SendWait("{DEL}");
        FormsSendKeys.SendWait(value);
    }

    private static void Invoke(AutomationElement el)
    {
        if (el.TryGetCurrentPattern(InvokePattern.Pattern, out var p) && p is InvokePattern ip)
        {
            ip.Invoke();
            return;
        }
        var r = el.Current.BoundingRectangle;
        if (r.Width <= 0 || r.Height <= 0) throw new InvalidOperationException("Нет координат для щелчка: " + el.Current.Name);
        var x = (int)(r.X + r.Width / 2);
        var y = (int)(r.Y + r.Height / 2);
        System.Windows.Forms.Cursor.Position = new DrawingPoint(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(40);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }

    [DllImport("user32.dll")] private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

    private static void SelectCombo(AutomationElement combo, string name)
    {
        combo.SetFocus();
        if (combo.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var p) && p is ExpandCollapsePattern exp)
            exp.Expand();
        Thread.Sleep(200);
        var item = combo.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name))
            ?? AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
        if (item != null && item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var sp) && sp is SelectionItemPattern sip)
            sip.Select();
        else
        {
            FormsSendKeys.SendWait("{HOME}");
            for (var i = 0; i < 8; i++)
            {
                if ((combo.GetCurrentPropertyValue(AutomationElement.NameProperty) as string)?.Contains(name) == true) break;
                FormsSendKeys.SendWait("{DOWN}");
                Thread.Sleep(80);
            }
            FormsSendKeys.SendWait("{ENTER}");
        }
        Thread.Sleep(200);
    }

    private static void SelectComboIndex(AutomationElement combo, int index)
    {
        combo.SetFocus();
        FormsSendKeys.SendWait("{HOME}");
        for (var i = 0; i < index; i++) FormsSendKeys.SendWait("{DOWN}");
        FormsSendKeys.SendWait("{ENTER}");
        Thread.Sleep(200);
    }

    private static void DismissDialogs()
    {
        for (var n = 0; n < 6; n++)
        {
            Thread.Sleep(250);
            var desktop = AutomationElement.RootElement;
            var dialogs = desktop.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
            var closed = false;
            foreach (AutomationElement d in dialogs)
            {
                var title = d.Current.Name ?? "";
                if (title is "Ошибка" or "Информация" or "Предупреждение")
                {
                    var ok = d.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ОК"))
                        ?? d.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "OK"));
                    if (ok != null) Invoke(ok);
                    closed = true;
                }
            }
            if (!closed) break;
        }
    }

    private static void Shot(AutomationElement win, string name)
    {
        var hwnd = new IntPtr(win.Current.NativeWindowHandle);
        ShowWindow(hwnd, 3);
        SetForegroundWindow(hwnd);
        Thread.Sleep(250);
        if (!PrintHwnd(hwnd, name)) ShotDesktop(name);
    }

    private static void ShotDialogOrMain(AutomationElement win, string name)
    {
        var desktop = AutomationElement.RootElement;
        var dialogs = desktop.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
        foreach (AutomationElement d in dialogs)
        {
            var title = d.Current.Name ?? "";
            if (title is "Ошибка" or "Информация" or "Предупреждение")
            {
                PrintHwnd(new IntPtr(win.Current.NativeWindowHandle), name + "-main");
                PrintHwnd(new IntPtr(d.Current.NativeWindowHandle), name);
                return;
            }
        }
        Shot(win, name);
    }

    private static void ShotDesktop(string name)
    {
        Thread.Sleep(200);
        var hwnd = app?.MainWindowHandle ?? IntPtr.Zero;
        if (hwnd != IntPtr.Zero && PrintHwnd(hwnd, name)) return;
        foreach (var p in Process.GetProcessesByName("Shop"))
        {
            if (p.MainWindowHandle != IntPtr.Zero && PrintHwnd(p.MainWindowHandle, name)) return;
        }
        var bounds = System.Windows.Forms.SystemInformation.VirtualScreen;
        using var bmp = new Bitmap(bounds.Width, bounds.Height);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(bounds.Location, DrawingPoint.Empty, bounds.Size);
        var path = IOPath.Combine(evidenceDir, name + ".png");
        bmp.Save(path, ImageFormat.Png);
        Console.WriteLine("saved desktop " + path);
    }

    private static bool PrintHwnd(IntPtr hwnd, string name)
    {
        if (hwnd == IntPtr.Zero) return false;
        if (!GetWindowRect(hwnd, out var rc)) return false;
        var w = Math.Max(1, rc.Right - rc.Left);
        var h = Math.Max(1, rc.Bottom - rc.Top);
        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        var hdc = g.GetHdc();
        var ok = PrintWindow(hwnd, hdc, 2);
        g.ReleaseHdc(hdc);
        if (!ok)
        {
            g.CopyFromScreen(rc.Left, rc.Top, 0, 0, new Size(w, h));
        }
        var path = IOPath.Combine(evidenceDir, name + ".png");
        bmp.Save(path, ImageFormat.Png);
        Console.WriteLine("saved " + path);
        return true;
    }
}

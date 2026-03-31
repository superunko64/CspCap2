using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Windows.Foundation;
using Windows.Graphics;
using Windows.Graphics.Capture;
using Windows.Graphics.DirectX;
using Windows.Graphics.DirectX.Direct3D11;
using Vortice.Direct3D;
using Vortice.Direct3D11;
using Vortice.DXGI;
using static Vortice.Direct3D11.D3D11;
using static Vortice.DXGI.DXGI;

ApplicationConfiguration.Initialize();
Application.Run(new CaptureForm());

class CaptureForm : Form
{
    readonly PictureBox picture = new() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.Black };
    readonly Button choose = new() { Text = "ウィンドウ選択", Dock = DockStyle.Top, Height = 40 };
    GraphicsCaptureSession? session;
    Direct3D11CaptureFramePool? pool;
    GraphicsCaptureItem? item;
    ID3D11Device d3d;
    ID3D11DeviceContext ctx;
    IDirect3DDevice winrtDevice;
    ID3D11Texture2D? staging;

    public CaptureForm()
    {
        Text = "キャプチャ表示ウィンドウ";
        Width = 1000;
        Height = 700;
        Controls.Add(picture);
        Controls.Add(choose);
        choose.Click += (_, _) => new PickerForm(this).Show(this);

        D3D11CreateDevice(null, DriverType.Hardware, DeviceCreationFlags.BgraSupport, null, out d3d, out ctx);
        CreateDirect3D11DeviceFromDXGIDevice(d3d.NativePointer, out var p);
        winrtDevice = (IDirect3DDevice)Marshal.GetObjectForIUnknown(p);
        Marshal.Release(p);
    }

    public void StartCapture(nint hwnd)
    {
        session?.Dispose();
        pool?.Dispose();
        item = CreateItemForWindow(hwnd);
        if (item == null) return;
        pool = Direct3D11CaptureFramePool.CreateFreeThreaded(winrtDevice, DirectXPixelFormat.B8G8R8A8UIntNormalized, 2, item.Size);
        session = pool.CreateCaptureSession(item);
        pool.FrameArrived += (_, _) => DrawFrame();
        session.StartCapture();
    }

    void DrawFrame()
    {
        if (pool == null) return;
        using var frame = pool.TryGetNextFrame();
        if (frame == null) return;
        if (frame.ContentSize.Width <= 0 || frame.ContentSize.Height <= 0) return;
        var tex = GetDXGIInterfaceFromObject<ID3D11Texture2D>(frame.Surface);
        if (staging == null || staging.Description.Width != tex.Description.Width || staging.Description.Height != tex.Description.Height)
        {
            staging?.Dispose();
            var d = tex.Description;
            d.Usage = ResourceUsage.Staging;
            d.BindFlags = BindFlags.None;
            d.CPUAccessFlags = CpuAccessFlags.Read;
            d.MiscFlags = ResourceOptionFlags.None;
            staging = d3d.CreateTexture2D(d);
            pool.Recreate(winrtDevice, DirectXPixelFormat.B8G8R8A8UIntNormalized, 2, frame.ContentSize);
        }
        ctx.CopyResource(staging!, tex);
        var box = ctx.Map(staging!, 0, MapMode.Read, Vortice.Direct3D11.MapFlags.None);
        var bmp = new Bitmap((int)staging!.Description.Width, (int)staging.Description.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        var bd = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat);
        var rowBytes = bmp.Width * 4;
        var rowBuffer = new byte[rowBytes];
        for (int y = 0; y < bmp.Height; y++)
        {
            var srcOffset = checked(y * (int)box.RowPitch);
            var srcRow = nint.Add(box.DataPointer, srcOffset);
            var dstRow = nint.Add(bd.Scan0, y * bd.Stride);
            Marshal.Copy(srcRow, rowBuffer, 0, rowBytes);
            Marshal.Copy(rowBuffer, 0, dstRow, rowBytes);
        }
        bmp.UnlockBits(bd);
        ctx.Unmap(staging!, 0);
        BeginInvoke(() =>
        {
            var old = picture.Image;
            picture.Image = bmp;
            old?.Dispose();
        });
    }

    static GraphicsCaptureItem CreateItemForWindow(nint hwnd)
    {
        var hstr = WindowsCreateString("Windows.Graphics.Capture.GraphicsCaptureItem", 45, out var hs);
        if (hstr != 0) throw new Exception("WindowsCreateString failed");
        try
        {
            var iid = typeof(IGraphicsCaptureItemInterop).GUID;
            RoGetActivationFactory(hs, ref iid, out var factoryPtr);
            var interop = (IGraphicsCaptureItemInterop)Marshal.GetObjectForIUnknown(factoryPtr);
            var itemGuid = typeof(GraphicsCaptureItem).GUID;
            interop.CreateForWindow(hwnd, ref itemGuid, out var itemPtr);
            Marshal.Release(factoryPtr);
            var gc = (GraphicsCaptureItem)Marshal.GetObjectForIUnknown(itemPtr);
            Marshal.Release(itemPtr);
            return gc;
        }
        finally { WindowsDeleteString(hs); }
    }

    static T GetDXGIInterfaceFromObject<T>(object obj)
    {
        var access = (IDirect3DDxgiInterfaceAccess)obj;
        var iid = typeof(T).GUID;
        access.GetInterface(ref iid, out var p);
        var t = (T)Marshal.GetObjectForIUnknown(p);
        Marshal.Release(p);
        return t;
    }

    [DllImport("d3d11.dll")]
    static extern int CreateDirect3D11DeviceFromDXGIDevice(nint dxgiDevice, out nint graphicsDevice);
    [DllImport("combase.dll")]
    static extern int RoGetActivationFactory(nint activatableClassId, ref Guid iid, out nint factory);
    [DllImport("combase.dll")]
    static extern int WindowsCreateString([MarshalAs(UnmanagedType.LPWStr)] string sourceString, int length, out nint hstring);
    [DllImport("combase.dll")]
    static extern int WindowsDeleteString(nint hstring);
}

class PickerForm : Form
{
    readonly ListBox list = new() { Dock = DockStyle.Fill };
    readonly CaptureForm owner;
    public PickerForm(CaptureForm owner)
    {
        this.owner = owner;
        Text = "ウィンドウ選択ウィンドウ";
        Width = 500;
        Height = 700;
        Controls.Add(list);
        list.DoubleClick += (_, _) =>
        {
            if (list.SelectedItem is WinItem w)
            {
                owner.StartCapture(w.Hwnd);
                Close();
            }
        };
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h)) return true;
            GetWindowText(h, out var title);
            if (!string.IsNullOrWhiteSpace(title)) list.Items.Add(new WinItem(h, title));
            return true;
        }, 0);
    }

    record WinItem(nint Hwnd, string Title)
    {
        public override string ToString() => $"0x{Hwnd.ToInt64():X}  {Title}";
    }

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);
    delegate bool EnumWindowsProc(nint hWnd, nint lParam);
    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(nint hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetWindowTextW(nint hWnd, char[] text, int maxCount);
    static void GetWindowText(nint hWnd, out string title)
    {
        var b = new char[512];
        var n = GetWindowTextW(hWnd, b, b.Length);
        title = new string(b, 0, n);
    }
}

[ComImport, Guid("3628e81b-3cac-4c60-b7f4-23ce0e0c3356"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IGraphicsCaptureItemInterop
{
    void CreateForWindow(nint window, ref Guid iid, out nint result);
    void CreateForMonitor(nint monitor, ref Guid iid, out nint result);
}

[ComImport, Guid("A9B3D012-3DF2-4EE3-B8D1-8695F457D3C1"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDirect3DDxgiInterfaceAccess
{
    void GetInterface(ref Guid iid, out nint p);
}

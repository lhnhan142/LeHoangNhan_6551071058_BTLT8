// ============================================================
//  Câu 1: Tiệm tạp hóa Nhà trường UTC2 – quản lý hàng hóa
//  1 file duy nhất: Model + DbContext + Form (dựng bằng code) + FastReport + ClosedXML
//  Tự tạo CSDL TiemTapHoa + 6 dòng mẫu ở lần chạy đầu (SQL Server LocalDB).
//  Nếu dùng SQL Server khác: sửa ConnStr bên dưới.
// ============================================================
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Diagnostics;
using ClosedXML.Excel;
using FastReport;
using FastReport.Export.PdfSimple;   // nếu dùng gói khác: FastReport.Export.Pdf + PDFExport
using FastReport.Format;
using Microsoft.EntityFrameworkCore;

ApplicationConfiguration.Initialize();
try
{
    using (var db = new TiemTapHoaContext())
    {
        db.Database.EnsureCreated();
        if (!db.HangHoas.Any())
        {
            db.HangHoas.AddRange(
                new HangHoa { TenHang = "Nước suối Aquafina 500ml", LoaiHang = "Đồ uống", DonVi = "Chai", GiaBan = 6000, SoLuongTon = 120, NgayNhap = new DateTime(2026, 10, 2), ConKinhDoanh = true },
                new HangHoa { TenHang = "Mì Hảo Hảo tôm chua cay", LoaiHang = "Thực phẩm khô", DonVi = "Gói", GiaBan = 4500, SoLuongTon = 200, NgayNhap = new DateTime(2026, 10, 2), ConKinhDoanh = true },
                new HangHoa { TenHang = "Dầu ăn Tường An 1L", LoaiHang = "Gia vị", DonVi = "Chai", GiaBan = 48000, SoLuongTon = 35, NgayNhap = new DateTime(2026, 10, 5), ConKinhDoanh = true },
                new HangHoa { TenHang = "Bột giặt OMO 800g", LoaiHang = "Đồ gia dụng", DonVi = "Túi", GiaBan = 52000, SoLuongTon = 40, NgayNhap = new DateTime(2026, 10, 5), ConKinhDoanh = true },
                new HangHoa { TenHang = "Sữa tươi Vinamilk 1L", LoaiHang = "Đồ uống", DonVi = "Hộp", GiaBan = 32000, SoLuongTon = 60, NgayNhap = new DateTime(2026, 10, 8), ConKinhDoanh = true },
                new HangHoa { TenHang = "Nước mắm Nam Ngư 500ml", LoaiHang = "Gia vị", DonVi = "Chai", GiaBan = 38000, SoLuongTon = 25, NgayNhap = new DateTime(2026, 10, 1), ConKinhDoanh = false });
            db.SaveChanges();
        }
    }
}
catch (Exception ex)
{
    MessageBox.Show("Không kết nối được SQL Server.\nHãy sửa ConnStr trong TiemTapHoaContext.\n\n" + ex.Message,
        "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
    return;
}
Application.Run(new FrmHangHoa());

// ====================== MODEL + DBCONTEXT ======================
[Table("HangHoa")]
public class HangHoa
{
    [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public int MaHang { get; set; }
    [Required, MaxLength(100)] public string TenHang { get; set; } = "";
    [MaxLength(50)] public string LoaiHang { get; set; }
    [MaxLength(20)] public string DonVi { get; set; }
    [Column(TypeName = "decimal(18,0)")] public decimal GiaBan { get; set; }
    public int SoLuongTon { get; set; }
    [Column(TypeName = "datetime")] public DateTime NgayNhap { get; set; }
    public bool ConKinhDoanh { get; set; }
}

public class TiemTapHoaContext : DbContext
{
    public const string ConnStr =
        @"Server=(localdb)\MSSQLLocalDB;Database=TiemTapHoa;Trusted_Connection=True;TrustServerCertificate=True";
    public DbSet<HangHoa> HangHoas { get; set; }
    protected override void OnConfiguring(DbContextOptionsBuilder o) => o.UseSqlServer(ConnStr);
}

// ====================== FORM ======================
public class FrmHangHoa : Form
{
    static readonly string[] Loais = { "Đồ uống", "Thực phẩm khô", "Gia vị", "Đồ gia dụng" };
    static readonly string[] DonVis = { "Chai", "Gói", "Hộp", "Túi" };

    TextBox txtTen, txtTim;
    ComboBox cboLoai, cboDonVi, cboLocLoai;
    NumericUpDown numGia, numTon;
    DateTimePicker dtpNgay;
    CheckBox chkKD, chkChiConBan;
    DataGridView dgv;
    ToolStripStatusLabel lblStatus;
    System.Windows.Forms.Timer timer;
    int _selectedId = 0;
    bool _loading = false;

    public FrmHangHoa()
    {
        Text = "Quản lý hàng hóa - Tạp hóa UTC2";
        Size = new Size(900, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F);
        BuildUi();
        timer = new System.Windows.Forms.Timer { Interval = 350 };
        timer.Tick += (s, e) => { timer.Stop(); LoadData(); };
        Load += (s, e) => LoadData();
    }

    // ---------- dựng giao diện ----------
    void BuildUi()
    {
        var gbInfo = new GroupBox { Text = "Thông tin hàng hóa", Location = new Point(12, 8), Size = new Size(520, 150) };
        var gbFind = new GroupBox { Text = "Tìm kiếm", Location = new Point(545, 8), Size = new Size(330, 150) };
        Controls.Add(gbInfo); Controls.Add(gbFind);

        gbInfo.Controls.Add(Lbl("Tên hàng:", 12, 27));
        txtTen = new TextBox { Location = new Point(85, 24), Width = 420 };
        gbInfo.Controls.Add(txtTen);

        gbInfo.Controls.Add(Lbl("Loại hàng:", 12, 58));
        cboLoai = new ComboBox { Location = new Point(85, 55), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
        cboLoai.Items.AddRange(Loais);
        gbInfo.Controls.Add(cboLoai);

        gbInfo.Controls.Add(Lbl("Đơn vị:", 12, 89));
        cboDonVi = new ComboBox { Location = new Point(85, 86), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
        cboDonVi.Items.AddRange(DonVis);
        gbInfo.Controls.Add(cboDonVi);

        gbInfo.Controls.Add(Lbl("Giá bán:", 12, 120));
        numGia = new NumericUpDown { Location = new Point(85, 117), Width = 150, Maximum = 1000000000, Increment = 500, ThousandsSeparator = true };
        gbInfo.Controls.Add(numGia);

        gbInfo.Controls.Add(Lbl("Số lượng tồn:", 255, 58));
        numTon = new NumericUpDown { Location = new Point(345, 55), Width = 160, Maximum = 1000000, Minimum = 0 };
        gbInfo.Controls.Add(numTon);

        gbInfo.Controls.Add(Lbl("Ngày nhập:", 255, 89));
        dtpNgay = new DateTimePicker { Location = new Point(345, 86), Width = 160, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
        gbInfo.Controls.Add(dtpNgay);

        chkKD = new CheckBox { Text = "Còn kinh doanh", Location = new Point(345, 120), AutoSize = true, Checked = true };
        gbInfo.Controls.Add(chkKD);

        txtTim = new TextBox { Location = new Point(14, 26), Width = 300 };
        txtTim.TextChanged += (s, e) => { timer.Stop(); timer.Start(); };   // chờ 350 ms sau khi ngừng gõ
        gbFind.Controls.Add(txtTim);

        cboLocLoai = new ComboBox { Location = new Point(14, 58), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cboLocLoai.Items.Add("Tất cả loại");
        cboLocLoai.Items.AddRange(Loais);
        cboLocLoai.SelectedIndex = 0;
        cboLocLoai.SelectedIndexChanged += (s, e) => LoadData();
        gbFind.Controls.Add(cboLocLoai);

        chkChiConBan = new CheckBox { Text = "Chỉ hiện hàng còn kinh doanh", Location = new Point(16, 92), AutoSize = true };
        chkChiConBan.CheckedChanged += (s, e) => LoadData();
        gbFind.Controls.Add(chkChiConBan);

        // ---- lưới ----
        dgv = new DataGridView
        {
            Location = new Point(12, 170), Size = new Size(863, 310),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AutoGenerateColumns = false, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false,
            RowHeadersVisible = true, BackgroundColor = SystemColors.AppWorkspace
        };
        dgv.Columns.Add(Col("MaHang", "Mã", 50));
        dgv.Columns.Add(Col("TenHang", "Tên hàng", 230));
        dgv.Columns.Add(Col("LoaiHang", "Loại hàng", 110));
        dgv.Columns.Add(Col("DonVi", "Đơn vị", 70));
        var cGia = Col("GiaBan", "Giá bán", 90); cGia.DefaultCellStyle.Format = "N0"; cGia.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        dgv.Columns.Add(cGia);
        var cTon = Col("SoLuongTon", "Tồn", 60); cTon.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        dgv.Columns.Add(cTon);
        var cNgay = Col("NgayNhap", "Ngày nhập", 100); cNgay.DefaultCellStyle.Format = "dd/MM/yyyy";
        dgv.Columns.Add(cNgay);
        dgv.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "ConKinhDoanh", HeaderText = "Kinh doanh", Width = 80 });
        dgv.CellClick += (s, e) => FillInputs();
        dgv.SelectionChanged += (s, e) => { if (!_loading) FillInputs(); };
        Controls.Add(dgv);

        // ---- nút ----
        int y = 492;
        AddBtn("Thêm", 12, y, 80, BtnThem);
        AddBtn("Sửa", 98, y, 80, BtnSua);
        AddBtn("Xóa", 184, y, 80, BtnXoa);
        AddBtn("Làm mới", 270, y, 80, BtnLamMoi);
        AddBtn("Xem bảng giá", 420, y, 100, BtnXemBangGia);
        AddBtn("Xuất PDF", 526, y, 80, BtnXuatPdf);
        AddBtn("Xuất Excel (lưới)", 612, y, 120, BtnExcelLuoi);
        AddBtn("Xuất Excel tồn kho", 738, y, 137, BtnExcelTonKho);

        var ss = new StatusStrip();
        lblStatus = new ToolStripStatusLabel();
        ss.Items.Add(lblStatus);
        Controls.Add(ss);
    }

    Label Lbl(string t, int x, int y) => new Label { Text = t, Location = new Point(x, y), AutoSize = true };
    DataGridViewTextBoxColumn Col(string prop, string head, int w) =>
        new DataGridViewTextBoxColumn { DataPropertyName = prop, HeaderText = head, Width = w };
    void AddBtn(string text, int x, int y, int w, EventHandler click)
    {
        var b = new Button { Text = text, Location = new Point(x, y), Size = new Size(w, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
        b.Click += click;
        Controls.Add(b);
    }

    // ---------- tải + lọc ----------
    void LoadData()
    {
        try
        {
            _loading = true;
            using var db = new TiemTapHoaContext();
            IQueryable<HangHoa> q = db.HangHoas.AsNoTracking();
            string kw = txtTim.Text.Trim();
            if (kw.Length > 0) q = q.Where(h => h.TenHang.Contains(kw));
            if (cboLocLoai.SelectedIndex > 0) { string loai = cboLocLoai.Text; q = q.Where(h => h.LoaiHang == loai); }
            if (chkChiConBan.Checked) q = q.Where(h => h.ConKinhDoanh);
            var list = q.OrderBy(h => h.MaHang).ToList();
            dgv.DataSource = list;
            dgv.ClearSelection();
            _selectedId = 0;
            lblStatus.Text = $"Tìm thấy {list.Count} mặt hàng";
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        finally { _loading = false; }
    }

    void FillInputs()
    {
        if (dgv.CurrentRow?.DataBoundItem is not HangHoa h || !dgv.CurrentRow.Selected) return;
        _selectedId = h.MaHang;
        txtTen.Text = h.TenHang;
        cboLoai.Text = h.LoaiHang;
        cboDonVi.Text = h.DonVi;
        numGia.Value = Math.Min(h.GiaBan, numGia.Maximum);
        numTon.Value = Math.Min(Math.Max(h.SoLuongTon, 0), numTon.Maximum);
        dtpNgay.Value = h.NgayNhap < dtpNgay.MinDate ? dtpNgay.MinDate : h.NgayNhap;
        chkKD.Checked = h.ConKinhDoanh;
    }

    bool Validate_()
    {
        if (string.IsNullOrWhiteSpace(txtTen.Text))
        { MessageBox.Show("Tên hàng không được để trống.", "Thiếu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning); txtTen.Focus(); return false; }
        if (numGia.Value <= 0)
        { MessageBox.Show("Giá bán phải lớn hơn 0.", "Sai dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning); numGia.Focus(); return false; }
        return true;
    }

    void ReadInputs(HangHoa h)
    {
        h.TenHang = txtTen.Text.Trim();
        h.LoaiHang = cboLoai.Text;
        h.DonVi = cboDonVi.Text;
        h.GiaBan = numGia.Value;
        h.SoLuongTon = (int)numTon.Value;
        h.NgayNhap = dtpNgay.Value.Date;
        h.ConKinhDoanh = chkKD.Checked;
    }

    void ClearInputs()
    {
        txtTen.Clear(); cboLoai.SelectedIndex = -1; cboDonVi.SelectedIndex = -1;
        numGia.Value = 0; numTon.Value = 0; dtpNgay.Value = DateTime.Today; chkKD.Checked = true;
        _selectedId = 0; txtTen.Focus();
    }

    // ---------- Thêm / Sửa / Xóa / Làm mới ----------
    void BtnThem(object s, EventArgs e)
    {
        if (!Validate_()) return;
        using (var db = new TiemTapHoaContext())
        {
            var h = new HangHoa(); ReadInputs(h);
            db.HangHoas.Add(h); db.SaveChanges();
        }
        LoadData(); ClearInputs();
        MessageBox.Show("Đã thêm hàng hóa.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnSua(object s, EventArgs e)
    {
        if (_selectedId == 0) { MessageBox.Show("Hãy chọn một dòng để sửa.", "Chưa chọn", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        if (!Validate_()) return;
        using (var db = new TiemTapHoaContext())
        {
            var h = db.HangHoas.Find(_selectedId);
            if (h == null) { MessageBox.Show("Không tìm thấy mặt hàng."); return; }
            ReadInputs(h); db.SaveChanges();
        }
        LoadData();
        MessageBox.Show("Đã cập nhật.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    void BtnXoa(object s, EventArgs e)
    {
        if (_selectedId == 0) { MessageBox.Show("Hãy chọn một dòng để xóa.", "Chưa chọn", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        if (MessageBox.Show($"Xóa mặt hàng \"{txtTen.Text}\"?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        using (var db = new TiemTapHoaContext())
        {
            var h = db.HangHoas.Find(_selectedId);
            if (h != null) { db.HangHoas.Remove(h); db.SaveChanges(); }
        }
        LoadData(); ClearInputs();
    }

    void BtnLamMoi(object s, EventArgs e)
    {
        ClearInputs();
        txtTim.Clear(); cboLocLoai.SelectedIndex = 0; chkChiConBan.Checked = false;
        LoadData();
    }

    // ---------- FastReport ----------
    DataTable LayHangConBan()
    {
        using var db = new TiemTapHoaContext();
        var list = db.HangHoas.AsNoTracking().Where(h => h.ConKinhDoanh).OrderBy(h => h.MaHang).ToList();
        var dt = new DataTable("HangHoa");
        dt.Columns.Add("TenHang", typeof(string));
        dt.Columns.Add("DonVi", typeof(string));
        dt.Columns.Add("GiaBan", typeof(decimal));
        foreach (var h in list) dt.Rows.Add(h.TenHang, h.DonVi, h.GiaBan);
        return dt;
    }

    // Dựng mẫu BangGia bằng code (không cần file .frx)
    Report BuildReport()
    {
        var rpt = new Report();
        rpt.RegisterData(LayHangConBan(), "HangHoa");
        rpt.GetDataSource("HangHoa").Enabled = true;
        rpt.Dictionary.Parameters.Add(new FastReport.Data.Parameter("NgayIn"));
        rpt.SetParameterValue("NgayIn", DateTime.Now.ToString("dd/MM/yyyy"));
        rpt.DoublePass = true;                       // để [TotalPages#] có giá trị

        var page = new ReportPage();
        rpt.Pages.Add(page);
        page.CreateUniqueName();

        float cm = FastReport.Utils.Units.Centimeters;
        float wSTT = 1.5f * cm, wTen = 10f * cm, wDV = 3f * cm, wGia = 4.5f * cm, total = wSTT + wTen + wDV + wGia;

        // Report Title
        page.ReportTitle = new ReportTitleBand();
        page.ReportTitle.CreateUniqueName();
        page.ReportTitle.Height = 1.8f * cm;
        var t1 = Cell(page.ReportTitle, "BẢNG GIÁ HÀNG HÓA – TẠP HÓA NHÀ TRƯỜNG UTC2", 0, 0, total, 1f * cm, HorzAlign.Center, 13, true, false);
        var t2 = Cell(page.ReportTitle, "Ngày in: [NgayIn]", 0, 1f * cm, total, 0.7f * cm, HorzAlign.Center, 10, false, false);

        // Page Header
        page.PageHeader = new PageHeaderBand();
        page.PageHeader.CreateUniqueName();
        page.PageHeader.Height = 0.8f * cm;
        Cell(page.PageHeader, "STT", 0, 0, wSTT, 0.8f * cm, HorzAlign.Center, 10, true, true);
        Cell(page.PageHeader, "Tên hàng", wSTT, 0, wTen, 0.8f * cm, HorzAlign.Center, 10, true, true);
        Cell(page.PageHeader, "Đơn vị", wSTT + wTen, 0, wDV, 0.8f * cm, HorzAlign.Center, 10, true, true);
        Cell(page.PageHeader, "Giá bán", wSTT + wTen + wDV, 0, wGia, 0.8f * cm, HorzAlign.Center, 10, true, true);

        // Data Band
        var band = new DataBand();
        page.Bands.Add(band);
        band.CreateUniqueName();
        band.DataSource = rpt.GetDataSource("HangHoa");
        band.Height = 0.7f * cm;
        Cell(band, "[Row#]", 0, 0, wSTT, 0.7f * cm, HorzAlign.Center, 10, false, true);
        Cell(band, "[HangHoa.TenHang]", wSTT, 0, wTen, 0.7f * cm, HorzAlign.Left, 10, false, true);
        Cell(band, "[HangHoa.DonVi]", wSTT + wTen, 0, wDV, 0.7f * cm, HorzAlign.Center, 10, false, true);
        var gia = Cell(band, "[HangHoa.GiaBan]", wSTT + wTen + wDV, 0, wGia, 0.7f * cm, HorzAlign.Right, 10, false, true);
        gia.Format = new NumberFormat { UseLocale = false, DecimalDigits = 0, DecimalSeparator = ",", GroupSeparator = "." };

        // Page Footer
        page.PageFooter = new PageFooterBand();
        page.PageFooter.CreateUniqueName();
        page.PageFooter.Height = 0.8f * cm;
        var foot = Cell(page.PageFooter, "Trang [Page#]/[TotalPages#]", 0, 0.1f * cm, total, 0.6f * cm, HorzAlign.Center, 9, false, false);
        foot.Border.Lines = BorderLines.Top;

        return rpt;
    }

    static TextObject Cell(Base parent, string text, float x, float y, float w, float h,
                           HorzAlign align, float size, bool bold, bool border)
    {
        var t = new TextObject();
        t.Parent = parent;
        t.CreateUniqueName();
        t.Bounds = new RectangleF(x, y, w, h);
        t.Text = text;
        t.HorzAlign = align;
        t.VertAlign = VertAlign.Center;
        t.Font = new Font("Tahoma", size, bold ? FontStyle.Bold : FontStyle.Regular);
        if (border) t.Border.Lines = BorderLines.All;
        return t;
    }

    void BtnXemBangGia(object s, EventArgs e)
    {
        try
        {
            var rpt = BuildReport();
            // Có bộ xem trước WinForms (FastReport.Net) thì gọi rpt.Show(); không thì xuất PDF tạm rồi mở
            var show = rpt.GetType().GetMethod("Show", Type.EmptyTypes);
            if (show != null) show.Invoke(rpt, null);
            else
            {
                rpt.Prepare();
                string tmp = Path.Combine(Path.GetTempPath(), "BangGia_xemtruoc.pdf");
                rpt.Export(new PDFSimpleExport(), tmp);
                Process.Start(new ProcessStartInfo(tmp) { UseShellExecute = true });
            }
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi report", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    void BtnXuatPdf(object s, EventArgs e)
    {
        using var sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = "BangGia.pdf" };
        if (sfd.ShowDialog() != DialogResult.OK) return;
        try
        {
            using var rpt = BuildReport();
            rpt.Prepare();
            rpt.Export(new PDFSimpleExport(), sfd.FileName);   // gói khác: new PDFExport()
            AskOpen(sfd.FileName, "PDF");
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi xuất PDF", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    void AskOpen(string path, string kind)
    {
        if (MessageBox.Show($"Đã xuất {kind} thành công.\nBạn có muốn mở file không?", "Hoàn tất",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    // ---------- Excel ----------
    // Xuất đúng các dòng đang hiển thị trên lưới; số ghi là số, ngày ghi là ngày
    void BtnExcelLuoi(object s, EventArgs e)
    {
        var rows = dgv.Rows.Cast<DataGridViewRow>().Select(r => r.DataBoundItem as HangHoa).Where(h => h != null).ToList();
        if (rows.Count == 0) { MessageBox.Show("Không có dữ liệu để xuất."); return; }
        using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "HangHoa_Luoi.xlsx" };
        if (sfd.ShowDialog() != DialogResult.OK) return;
        try
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("HangHoa");
            string[] heads = { "Mã", "Tên hàng", "Loại hàng", "Đơn vị", "Giá bán", "Tồn", "Ngày nhập", "Kinh doanh" };
            for (int i = 0; i < heads.Length; i++) ws.Cell(1, i + 1).Value = heads[i];
            int r = 2;
            foreach (var h in rows)
            {
                ws.Cell(r, 1).Value = h.MaHang;
                ws.Cell(r, 2).Value = h.TenHang;
                ws.Cell(r, 3).Value = h.LoaiHang;
                ws.Cell(r, 4).Value = h.DonVi;
                ws.Cell(r, 5).Value = h.GiaBan;
                ws.Cell(r, 6).Value = h.SoLuongTon;
                ws.Cell(r, 7).Value = h.NgayNhap;
                ws.Cell(r, 8).Value = h.ConKinhDoanh ? "Còn bán" : "Ngừng bán";
                r++;
            }
            StyleSheet(ws, r - 1);
            wb.SaveAs(sfd.FileName);
            AskOpen(sfd.FileName, "Excel");
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi xuất Excel", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    // Xuất từ truy vấn EF Core (toàn bộ hàng) + cột Giá trị tồn bằng công thức + dòng tổng
    void BtnExcelTonKho(object s, EventArgs e)
    {
        using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "TonKho.xlsx" };
        if (sfd.ShowDialog() != DialogResult.OK) return;
        try
        {
            List<HangHoa> list;
            using (var db = new TiemTapHoaContext())
                list = db.HangHoas.AsNoTracking().OrderBy(h => h.LoaiHang).ThenBy(h => h.TenHang).ToList();

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("TonKho");
            string[] heads = { "Mã", "Tên hàng", "Loại hàng", "Đơn vị", "Giá bán", "Tồn", "Ngày nhập", "Kinh doanh", "Giá trị tồn" };
            for (int i = 0; i < heads.Length; i++) ws.Cell(1, i + 1).Value = heads[i];
            int r = 2;
            foreach (var h in list)
            {
                ws.Cell(r, 1).Value = h.MaHang;
                ws.Cell(r, 2).Value = h.TenHang;
                ws.Cell(r, 3).Value = h.LoaiHang;
                ws.Cell(r, 4).Value = h.DonVi;
                ws.Cell(r, 5).Value = h.GiaBan;
                ws.Cell(r, 6).Value = h.SoLuongTon;
                ws.Cell(r, 7).Value = h.NgayNhap;
                ws.Cell(r, 8).Value = h.ConKinhDoanh ? "Còn bán" : "Ngừng bán";
                ws.Cell(r, 9).FormulaA1 = $"E{r}*F{r}";
                r++;
            }
            int last = r - 1;
            ws.Cell(r, 8).Value = "Tổng:";
            ws.Cell(r, 8).Style.Font.Bold = true;
            ws.Cell(r, 6).FormulaA1 = $"SUM(F2:F{last})";
            ws.Cell(r, 9).FormulaA1 = $"SUM(I2:I{last})";
            ws.Range(r, 6, r, 9).Style.Font.Bold = true;
            StyleSheet(ws, r);
            ws.Range(2, 9, r, 9).Style.NumberFormat.Format = "#,##0";
            wb.SaveAs(sfd.FileName);
            AskOpen(sfd.FileName, "Excel");
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi xuất Excel", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    static void StyleSheet(IXLWorksheet ws, int lastRow)
    {
        var head = ws.Range(1, 1, 1, ws.LastColumnUsed().ColumnNumber());
        head.Style.Font.Bold = true;
        head.Style.Fill.BackgroundColor = XLColor.LightGray;
        head.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range(2, 5, lastRow, 5).Style.NumberFormat.Format = "#,##0";          // Giá bán
        ws.Range(2, 6, lastRow, 6).Style.NumberFormat.Format = "#,##0";          // Tồn
        ws.Range(2, 7, lastRow, 7).Style.NumberFormat.Format = "dd/MM/yyyy";     // Ngày nhập
        ws.SheetView.FreezeRows(1);
        ws.Columns().AdjustToContents();
    }
}

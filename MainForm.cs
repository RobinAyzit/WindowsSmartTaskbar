using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace WindowsSmartTaskbar
{
    public partial class MainForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private List<ProgramItem> programs = new List<ProgramItem>();
        private List<Category> categories = new List<Category>();
        private string currentCategory = "All programs";
        private static string AppDataFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WindowsSmartTaskbar");
        private static string ConfigFile => Path.Combine(AppDataFolder, "programs.json");
        private static string CategoryFile => Path.Combine(AppDataFolder, "categories.json");
        private static string SettingsFile => Path.Combine(AppDataFolder, "settings.json");
        private const int MaxPrograms = 20;
        private const string DefaultCategory = "All programs";
        private NotifyIcon? notifyIcon;
        private ContextMenuStrip? contextMenu;
        private ContextMenuStrip? leftClickMenu;
        private System.Windows.Forms.Timer? clickTimer;
        private DateTime lastDoubleClickTime = DateTime.MinValue;

        // UI controls
        private Panel? programPanel;
        private Label? statusLabel;
        private Label? copyrightLabel;
        private ComboBox? categoryComboBox;
        private Label? titleLabel;
        private Button? addButton;
        private Button? removeButton;
        private Button? editButton;
        private Button? categoryButton;
        private Button? removeCategoryButton;
        private Button? editCategoryButton;
        private Label? categoryLabel;

        // Selection tracking
        private HashSet<int> selectedIndices = new HashSet<int>();
        
        // Drag and Drop
        private ProgramItem? draggedItem;
        private Point dragStartPoint;

        public List<ProgramItem> GetPrograms() => programs;
        public List<Category> GetCategories() => categories;
        
        public void SetProgramCategory(ProgramItem program, string category)
        {
            program.Category = category;
            SavePrograms();
            RefreshProgramList();
        }

        public void MoveProgram(ProgramItem source, ProgramItem target)
        {
            if (source == target) return;

            int oldIndex = programs.IndexOf(source);
            int newIndex = programs.IndexOf(target);

            if (oldIndex >= 0 && newIndex >= 0)
            {
                programs.RemoveAt(oldIndex);
                programs.Insert(newIndex, source);
                SavePrograms();
                RefreshProgramList();
            }
        }

        // Localization
        private string currentLanguage = "en";
        private string currentTheme = "dark";
        private static readonly Dictionary<string, Dictionary<string, string>> Strings = new Dictionary<string, Dictionary<string, string>>
        {
            ["sv"] = new Dictionary<string, string>
            {
                ["title"] = "WindowsSmartTaskbar",
                ["category"] = "Kategori:",
                ["addProgram"] = "Lägg till program",
                ["remove"] = "Ta bort",
                ["editName"] = "Redigera namn",
                ["addCategory"] = "Lägg till kategori",
                ["removeCategory"] = "Ta bort kategori",
                ["allPrograms"] = "Alla program",
                ["showPrograms"] = "Visa program",
                ["exit"] = "Avsluta",
                ["startWithWindows"] = "Starta med Windows",
                ["settings"] = "Inställningar",
                ["language"] = "Språk",
                ["resetAll"] = "Radera allt",
                ["resetConfirm"] = "Är du säker på att du vill radera alla program och kategorier?",
                ["resetTitle"] = "Bekräfta radering",
                ["resetDone"] = "Alla program och kategorier har raderats.",
                ["limitReached"] = "Du kan bara lägga till max {0} program.",
                ["selectFile"] = "Välj ett program eller genväg att lägga till",
                ["fileFilter"] = "Program och genvägar (*.exe;*.lnk)|*.exe;*.lnk|Programfiler (*.exe)|*.exe|Genvägar (*.lnk)|*.lnk|Alla filer (*.*)|*.*",
                ["noSelection"] = "Inget valt",
                ["selectProgramRemove"] = "Välj minst ett program att ta bort.",
                ["selectProgramEdit"] = "Välj ett program att redigera.",
                ["editNameTitle"] = "Redigera programnamn",
                ["newName"] = "Nytt namn:",
                ["ok"] = "OK",
                ["cancel"] = "Avbryt",
                ["addCategoryTitle"] = "Lägg till kategori",
                ["categoryName"] = "Kategorinamn:",
                ["categoryAdded"] = "Kategorin '{0}' har lagts till!",
                ["categoryExists"] = "Kategorin finns redan.",
                ["categoryRemoved"] = "Kategorin '{0}' har tagits bort!",
                ["cannotRemoveDefault"] = "Kan inte ta bort standardkategorin 'Alla program'.",
                ["programsInCategory"] = "Det finns {0} program i kategorin '{1}'.\n\nVill du flytta dessa program till 'Alla program'?",
                ["error"] = "Fel",
                ["programCount"] = "Program: {0}/{1}",
                ["couldNotLoad"] = "Kunde inte ladda: {0}",
                ["couldNotSave"] = "Kunde inte spara: {0}",
                ["couldNotStart"] = "Kunde inte starta: {0}",
                ["selectCategoryRemove"] = "Välj en kategori att ta bort.",
                ["editCategory"] = "Redigera kategori",
                ["editCategoryTitle"] = "Redigera kategorinamn",
                ["selectCategoryEdit"] = "Välj en kategori att redigera.",
                ["cannotEditDefault"] = "Kan inte redigera standardkategorin.",
                ["categoryRenamed"] = "Kategorin '{0}' har döpts om till '{1}'!",
                ["theme"] = "Tema",
                ["dark"] = "Mörkt",
                ["light"] = "Ljust",
                ["moveToCategory"] = "Flytta till kategori",
                ["contextDelete"] = "Radera",
                ["contextRename"] = "Redigera namn",
            },
            ["en"] = new Dictionary<string, string>
            {
                ["title"] = "WindowsSmartTaskbar",
                ["category"] = "Category:",
                ["addProgram"] = "Add program",
                ["remove"] = "Remove",
                ["editName"] = "Edit name",
                ["addCategory"] = "Add category",
                ["removeCategory"] = "Remove category",
                ["allPrograms"] = "All programs",
                ["showPrograms"] = "Show programs",
                ["exit"] = "Exit",
                ["startWithWindows"] = "Start with Windows",
                ["settings"] = "Settings",
                ["language"] = "Language",
                ["resetAll"] = "Reset all",
                ["resetConfirm"] = "Are you sure you want to delete all programs and categories?",
                ["resetTitle"] = "Confirm reset",
                ["resetDone"] = "All programs and categories have been deleted.",
                ["limitReached"] = "You can only add up to {0} programs.",
                ["selectFile"] = "Select a program or shortcut to add",
                ["fileFilter"] = "Programs and shortcuts (*.exe;*.lnk)|*.exe;*.lnk|Program files (*.exe)|*.exe|Shortcuts (*.lnk)|*.lnk|All files (*.*)|*.*",
                ["noSelection"] = "Nothing selected",
                ["selectProgramRemove"] = "Select at least one program to remove.",
                ["selectProgramEdit"] = "Select a program to edit.",
                ["editNameTitle"] = "Edit program name",
                ["newName"] = "New name:",
                ["ok"] = "OK",
                ["cancel"] = "Cancel",
                ["addCategoryTitle"] = "Add category",
                ["categoryName"] = "Category name:",
                ["categoryAdded"] = "Category '{0}' has been added!",
                ["categoryExists"] = "Category already exists.",
                ["categoryRemoved"] = "Category '{0}' has been removed!",
                ["cannotRemoveDefault"] = "Cannot remove the default category.",
                ["programsInCategory"] = "There are {0} programs in category '{1}'.\n\nDo you want to move them to the default category?",
                ["error"] = "Error",
                ["programCount"] = "Programs: {0}/{1}",
                ["couldNotLoad"] = "Could not load: {0}",
                ["couldNotSave"] = "Could not save: {0}",
                ["couldNotStart"] = "Could not start: {0}",
                ["selectCategoryRemove"] = "Select a category to remove.",
                ["editCategory"] = "Edit category",
                ["editCategoryTitle"] = "Edit category name",
                ["selectCategoryEdit"] = "Select a category to edit.",
                ["cannotEditDefault"] = "Cannot edit the default category.",
                ["categoryRenamed"] = "Category '{0}' has been renamed to '{1}'!",
                ["theme"] = "Theme",
                ["dark"] = "Dark",
                ["light"] = "Light",
                ["moveToCategory"] = "Move to category",
                ["contextDelete"] = "Delete",
                ["contextRename"] = "Rename",
            },
            ["tr"] = new Dictionary<string, string>
            {
                ["title"] = "WindowsSmartTaskbar",
                ["category"] = "Kategori:",
                ["addProgram"] = "Program ekle",
                ["remove"] = "Kaldır",
                ["editName"] = "Adı düzenle",
                ["addCategory"] = "Kategori ekle",
                ["removeCategory"] = "Kategoriyi kaldır",
                ["allPrograms"] = "Tüm programlar",
                ["showPrograms"] = "Programları göster",
                ["exit"] = "Çıkış",
                ["startWithWindows"] = "Windows ile başlat",
                ["settings"] = "Ayarlar",
                ["language"] = "Dil",
                ["resetAll"] = "Tümünü sıfırla",
                ["resetConfirm"] = "Tüm programları ve kategorileri silmek istediğinizden emin misiniz?",
                ["resetTitle"] = "Silmeyi onayla",
                ["resetDone"] = "Tüm programlar ve kategoriler silindi.",
                ["limitReached"] = "En fazla {0} program ekleyebilirsiniz.",
                ["selectFile"] = "Eklenecek bir program veya kısayol seçin",
                ["fileFilter"] = "Programlar ve kısayollar (*.exe;*.lnk)|*.exe;*.lnk|Program dosyaları (*.exe)|*.exe|Kısayollar (*.lnk)|*.lnk|Tüm dosyalar (*.*)|*.*",
                ["noSelection"] = "Seçim yok",
                ["selectProgramRemove"] = "Kaldırmak için en az bir program seçin.",
                ["selectProgramEdit"] = "Düzenlemek için bir program seçin.",
                ["editNameTitle"] = "Program adını düzenle",
                ["newName"] = "Yeni ad:",
                ["ok"] = "Tamam",
                ["cancel"] = "İptal",
                ["addCategoryTitle"] = "Kategori ekle",
                ["categoryName"] = "Kategori adı:",
                ["categoryAdded"] = "'{0}' kategorisi eklendi!",
                ["categoryExists"] = "Kategori zaten mevcut.",
                ["categoryRemoved"] = "'{0}' kategorisi kaldırıldı!",
                ["cannotRemoveDefault"] = "Varsayılan kategori kaldırılamaz.",
                ["programsInCategory"] = "'{1}' kategorisinde {0} program var.\n\nBunları varsayılan kategoriye taşımak ister misiniz?",
                ["error"] = "Hata",
                ["programCount"] = "Program: {0}/{1}",
                ["couldNotLoad"] = "Yüklenemedi: {0}",
                ["couldNotSave"] = "Kaydedilemedi: {0}",
                ["couldNotStart"] = "Başlatılamadı: {0}",
                ["selectCategoryRemove"] = "Kaldırmak için bir kategori seçin.",
                ["editCategory"] = "Kategoriyi düzenle",
                ["editCategoryTitle"] = "Kategori adını düzenle",
                ["selectCategoryEdit"] = "Düzenlemek için bir kategori seçin.",
                ["cannotEditDefault"] = "Varsayılan kategori düzenlenemez.",
                ["categoryRenamed"] = "'{0}' kategorisinin adı '{1}' olarak değiştirildi!",
                ["theme"] = "Tema",
                ["dark"] = "Karanlık",
                ["light"] = "Aydınlık",
                ["moveToCategory"] = "Kategoriye taşı",
                ["contextDelete"] = "Sil",
                ["contextRename"] = "Adını değiştir",
            }
        };

        private string T(string key)
        {
            if (Strings.TryGetValue(currentLanguage, out var dict) && dict.TryGetValue(key, out var val))
                return val;
            if (Strings["en"].TryGetValue(key, out var fallback))
                return fallback;
            return key;
        }

        public MainForm()
        {
            EnsureDataFolder();
            LoadSettings();
            InitializeComponent();
            LoadCategories();
            LoadPrograms();
            ApplyTheme();
            this.Icon = GenerateAppIcon();
            SetupNotifyIcon();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
        }

        private void EnsureDataFolder()
        {
            try
            {
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                }
            }
            catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "WindowsSmartTaskbar";
            this.Size = new Size(420, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = currentTheme == "dark" ? Color.FromArgb(32, 32, 32) : Color.FromArgb(245, 245, 250);

            // Huvudpanel
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // Titellabel
            titleLabel = new Label
            {
                Text = T("title"),
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 40,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(0, 120, 215)
            };

            // Kategoriväljare
            var categoryPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 35,
                FlowDirection = FlowDirection.LeftToRight
            };

            categoryLabel = new Label
            {
                Text = T("category"),
                Width = 65,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9)
            };

            categoryComboBox = new ComboBox
            {
                Width = 220,
                Height = 30,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9)
            };
            categoryComboBox.SelectedIndexChanged += CategoryComboBox_SelectedIndexChanged;

            categoryPanel.Controls.Add(categoryLabel);
            categoryPanel.Controls.Add(categoryComboBox);

            // Knapppanel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 135,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 5, 0, 5)
            };

            addButton = CreateButton(T("addProgram"), Color.FromArgb(0, 120, 215), Color.White, 120);
            addButton.Click += AddButton_Click;

            removeButton = CreateButton(T("remove"), Color.FromArgb(220, 53, 69), Color.White, 80);
            removeButton.Click += RemoveButton_Click;

            editButton = CreateButton(T("editName"), Color.FromArgb(255, 193, 7), Color.Black, 105);
            editButton.Click += EditButton_Click;

            categoryButton = CreateButton(T("addCategory"), Color.FromArgb(40, 167, 69), Color.White, 115);
            categoryButton.Click += CategoryButton_Click;

            removeCategoryButton = CreateButton(T("removeCategory"), Color.FromArgb(220, 53, 69), Color.White, 115);
            removeCategoryButton.Click += RemoveCategoryButton_Click;

            editCategoryButton = CreateButton(T("editCategory"), Color.FromArgb(255, 193, 7), Color.Black, 115);
            editCategoryButton.Click += EditCategoryButton_Click;

            buttonPanel.Controls.Add(addButton);
            buttonPanel.Controls.Add(removeButton);
            buttonPanel.Controls.Add(editButton);
            buttonPanel.Controls.Add(categoryButton);
            buttonPanel.Controls.Add(removeCategoryButton);
            buttonPanel.Controls.Add(editCategoryButton);

            // Programpanel (scrollbar panel istället för ListBox)
            programPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = currentTheme == "dark" ? Color.FromArgb(45, 45, 45) : Color.White,
                BorderStyle = currentTheme == "dark" ? BorderStyle.None : BorderStyle.FixedSingle,
                Padding = new Padding(0)
            };

            // Statuslabel with copyright
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
    Height = 60
            };

            statusLabel = new Label
            {
                Text = string.Format(T("programCount"), 0, MaxPrograms),
                Dock = DockStyle.Left,
                Width = 150,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };

            copyrightLabel = new Label
            {
                Text = "Created 2026 by © nRn World",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Padding = new Padding(0, 0, 10, 0),
                AutoSize = false
            };

            bottomPanel.Controls.Add(copyrightLabel);
            bottomPanel.Controls.Add(statusLabel);

            // OBS: Ordning är viktig för WinForms docking
            // Lägg till i omvänd ordning: Bottom först, sedan Fill, sedan Top
            mainPanel.Controls.Add(programPanel);      // Fill - ska vara först
            mainPanel.Controls.Add(bottomPanel);             // Bottom
            mainPanel.Controls.Add(buttonPanel);        // Top
            mainPanel.Controls.Add(categoryPanel);      // Top
            mainPanel.Controls.Add(titleLabel);         // Top

            this.Controls.Add(mainPanel);
        }

        private Button CreateButton(string text, Color backColor, Color foreColor, int width)
        {
            var btn = new Button
            {
                Text = text,
                Width = width,
                Height = 35,
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void SetupNotifyIcon()
        {
            notifyIcon = new NotifyIcon
            {
                Text = "WindowsSmartTaskbar",
                Visible = true
            };

            notifyIcon.Icon = GenerateAppIcon();
            contextMenu = new ContextMenuStrip();

            var showMenuItem = new ToolStripMenuItem(T("showPrograms"));
            showMenuItem.Click += (s, e) => ShowProgramList();
            contextMenu.Items.Add(showMenuItem);

            contextMenu.Items.Add(new ToolStripSeparator());

            var autostartMenuItem = new ToolStripMenuItem(T("startWithWindows"))
            {
                CheckOnClick = true,
                Checked = IsAutostartEnabled()
            };
            autostartMenuItem.Click += (s, e) => SetAutostart(autostartMenuItem.Checked);
            contextMenu.Items.Add(autostartMenuItem);

            // Settings submenu
            var settingsMenu = new ToolStripMenuItem(T("settings"));

            // Language submenu
            var langMenu = new ToolStripMenuItem(T("language"));
            var svItem = new ToolStripMenuItem("Svenska") { Tag = "sv", Checked = currentLanguage == "sv" };
            var enItem = new ToolStripMenuItem("English") { Tag = "en", Checked = currentLanguage == "en" };
            var trItem = new ToolStripMenuItem("Türkçe") { Tag = "tr", Checked = currentLanguage == "tr" };

            svItem.Click += (s, e) => ChangeLanguage("sv");
            enItem.Click += (s, e) => ChangeLanguage("en");
            trItem.Click += (s, e) => ChangeLanguage("tr");

            langMenu.DropDownItems.AddRange(new ToolStripItem[] { svItem, enItem, trItem });
            settingsMenu.DropDownItems.Add(langMenu);

            // Theme submenu
            var themeMenu = new ToolStripMenuItem(T("theme"));
            var darkItem = new ToolStripMenuItem(T("dark")) { Tag = "dark", Checked = currentTheme == "dark" };
            var lightItem = new ToolStripMenuItem(T("light")) { Tag = "light", Checked = currentTheme == "light" };

            darkItem.Click += (s, e) => ChangeTheme("dark");
            lightItem.Click += (s, e) => ChangeTheme("light");

            themeMenu.DropDownItems.AddRange(new ToolStripItem[] { darkItem, lightItem });
            settingsMenu.DropDownItems.Add(themeMenu);

            // Reset all
            var resetItem = new ToolStripMenuItem(T("resetAll"));
            resetItem.Click += ResetAllButton_Click;
            settingsMenu.DropDownItems.Add(resetItem);

            contextMenu.Items.Add(settingsMenu);

            contextMenu.Items.Add(new ToolStripSeparator());

            var exitMenuItem = new ToolStripMenuItem(T("exit"));
            exitMenuItem.Click += (s, e) => Application.Exit();
            contextMenu.Items.Add(exitMenuItem);

            notifyIcon.ContextMenuStrip = contextMenu;

            clickTimer = new System.Windows.Forms.Timer();
            clickTimer.Interval = SystemInformation.DoubleClickTime;
            clickTimer.Tick += (s, e) =>
            {
                clickTimer.Stop();
                BuildQuickLaunchMenu();
                SetForegroundWindow(this.Handle); // Ensure focus for menu
                leftClickMenu?.Show(Cursor.Position);
            };

            notifyIcon.MouseClick += (s, e) => 
            { 
                if (e.Button == MouseButtons.Left) 
                {
                    clickTimer.Start();
                }
            };
            
            notifyIcon.MouseDoubleClick += (s, e) => 
            { 
                if (e.Button == MouseButtons.Left) 
                {
                    clickTimer.Stop();
                    ShowProgramList(); 
                }
            };
        }

        private void BuildQuickLaunchMenu()
        {
            if (leftClickMenu == null)
            {
                leftClickMenu = new ContextMenuStrip();
            }
            else
            {
                leftClickMenu.Items.Clear();
            }

            foreach (var cat in categories)
            {
                var displayName = cat.Name == DefaultCategory ? T("allPrograms") : cat.Name;
                var catItem = new ToolStripMenuItem(displayName);
                
                var categoryPrograms = programs.Where(p => p.Category == cat.Name).ToList();
                if (cat.Name == DefaultCategory)
                {
                    categoryPrograms = programs.ToList();
                }

                foreach (var prog in categoryPrograms)
                {
                    var progItem = new ToolStripMenuItem(prog.Name);
                    if (prog.Icon != null)
                    {
                        try { progItem.Image = prog.Icon.ToBitmap(); } catch { }
                    }
                    progItem.Click += (s, e) => prog.Start();
                    catItem.DropDownItems.Add(progItem);
                }

                if (catItem.DropDownItems.Count > 0 || cat.Name == DefaultCategory)
                {
                    leftClickMenu.Items.Add(catItem);
                }
            }

            if (leftClickMenu.Items.Count == 0)
            {
                leftClickMenu.Items.Add(new ToolStripMenuItem(T("noSelection")) { Enabled = false });
            }
        }

        private Icon GenerateAppIcon()
        {
            // Skapa en modern ikon programmatiskt
            int size = 256;
            using (Bitmap bmp = new Bitmap(size, size))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.Clear(Color.Transparent);

                    // Bakgrund: Blå gradient cirkel
                    Rectangle rect = new Rectangle(10, 10, size - 20, size - 20);
                    using (var path = new System.Drawing.Drawing2D.GraphicsPath())
                    {
                        path.AddEllipse(rect);
                        using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(rect, 
                            Color.FromArgb(0, 120, 215), // Windows Blue
                            Color.FromArgb(0, 60, 140),  // Darker Blue
                            45f))
                        {
                            g.FillPath(brush, path);
                        }
                    }

                    // Symbol: Vita streck (launcher-koncept)
                    using (Pen pen = new Pen(Color.White, 16))
                    {
                        pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                        pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;

                        int centerX = size / 2;
                        int centerY = size / 2;
                        
                        // Tre horisontella streck
                        g.DrawLine(pen, centerX - 60, centerY - 45, centerX + 60, centerY - 45);
                        g.DrawLine(pen, centerX - 60, centerY, centerX + 30, centerY);
                        g.DrawLine(pen, centerX - 60, centerY + 45, centerX + 10, centerY + 45);
                        
                        // En liten accent-prick
                        g.FillEllipse(Brushes.White, centerX + 40, centerY + 35, 20, 20);
                    }
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }

        private void ChangeLanguage(string lang)
        {
            currentLanguage = lang;
            SaveSettings();

            // Update all UI text
            if (titleLabel != null) titleLabel.Text = T("title");
            if (categoryLabel != null) categoryLabel.Text = T("category");
            if (addButton != null) addButton.Text = T("addProgram");
            if (removeButton != null) removeButton.Text = T("remove");
            if (editButton != null) editButton.Text = T("editName");
            if (categoryButton != null) categoryButton.Text = T("addCategory");
            if (removeCategoryButton != null) removeCategoryButton.Text = T("removeCategory");
            UpdateStatus();
            LoadCategories();

            // Rebuild the tray menu
            notifyIcon?.Dispose();
            notifyIcon = null;
            contextMenu = null;
            SetupNotifyIcon();
        }

        private void ResetAllButton_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show(T("resetConfirm"), T("resetTitle"),
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                programs.Clear();
                categories.Clear();
                categories.Add(new Category(DefaultCategory));
                SavePrograms();
                SaveCategories();
                LoadCategories();
                RefreshProgramList();
                MessageBox.Show(T("resetDone"), T("title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ChangeTheme(string theme)
        {
            currentTheme = theme;
            SaveSettings();
            ApplyTheme();

            // Refresh tray menu to update checkmarks
            notifyIcon?.Dispose();
            notifyIcon = null;
            contextMenu = null;
            SetupNotifyIcon();
        }

        private void ApplyTheme()
        {
            Color bgColor = currentTheme == "dark" ? Color.FromArgb(32, 32, 32) : Color.FromArgb(245, 245, 250);
            Color panelBg = currentTheme == "dark" ? Color.FromArgb(45, 45, 45) : Color.White;
            Color textColor = currentTheme == "dark" ? Color.White : Color.FromArgb(30, 30, 30);
            Color secondaryTextColor = currentTheme == "dark" ? Color.FromArgb(180, 180, 180) : Color.Gray;
            Color borderColor = currentTheme == "dark" ? Color.FromArgb(60, 60, 60) : Color.FromArgb(230, 230, 230);

            this.BackColor = bgColor;
            
            if (programPanel != null)
            {
                programPanel.BackColor = panelBg;
                programPanel.BorderStyle = currentTheme == "dark" ? BorderStyle.None : BorderStyle.FixedSingle;
            }

            if (categoryLabel != null) categoryLabel.ForeColor = textColor;
            if (statusLabel != null) statusLabel.ForeColor = secondaryTextColor;
            if (copyrightLabel != null) copyrightLabel.ForeColor = secondaryTextColor;

            // Update nested controls if any
            RefreshProgramList();
        }

        private bool autostartEnabled = true; // Standard värde

        private void LoadSettings()
        {
            try
            {
                bool isFirstRun = false;
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    if (settings != null)
                    {
                        currentLanguage = settings.Language ?? "en";
                        currentTheme = settings.Theme ?? "dark";
                        autostartEnabled = settings.Autostart ?? true; // Standard true
                    }
                    else
                    {
                        isFirstRun = true;
                    }
                }
                else
                {
                    isFirstRun = true;
                }

                // Om det är första körningen, spara standardinställningarna
                if (isFirstRun)
                {
                    autostartEnabled = true;
                    SaveSettings();
                    SetAutostart(true);
                }
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                var settings = new AppSettings 
                { 
                    Language = currentLanguage,
                    Theme = currentTheme,
                    Autostart = autostartEnabled
                };
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFile, json);
            }
            catch { }
        }

        private bool IsAutostartEnabled()
        {
            // Returnera den sparade inställningen från AppSettings
            return autostartEnabled;
        }

        private void SetAutostart(bool enable)
        {
            // Uppdatera den sparade inställningen
            autostartEnabled = enable;
            SaveSettings();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key == null) return;
                    if (enable)
                    {
                        var currentPath = Environment.ProcessPath ?? Application.ExecutablePath;
                        key.SetValue("WindowsSmartTaskbar", $"\"{currentPath}\"");
                    }
                    else
                    {
                        key.DeleteValue("WindowsSmartTaskbar", false);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(T("couldNotSave"), ex.Message), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowProgramList()
        {
            clickTimer?.Stop();
            leftClickMenu?.Hide();
            leftClickMenu?.Close();

            this.ShowInTaskbar = true;
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;
            
            this.Show();
            this.Activate();
            this.BringToFront();
            SetForegroundWindow(this.Handle);
        }

        // ========== Program List Display (Panel-based, no OwnerDraw) ==========

        private void RefreshProgramList()
        {
            if (programPanel == null) return;

            programPanel.SuspendLayout();
            programPanel.Controls.Clear();
            selectedIndices.Clear();

            // Filter programs
            var filtered = currentCategory == DefaultCategory
                ? programs.ToList()
                : programs.Where(p => p.Category == currentCategory).ToList();

            // Add program panels in reverse so Dock.Top stacks them correctly
            for (int i = filtered.Count - 1; i >= 0; i--)
            {
                var panel = CreateProgramRow(filtered[i], i);
                programPanel.Controls.Add(panel);
            }

            programPanel.ResumeLayout();
            UpdateStatus();
        }

        private Panel CreateProgramRow(ProgramItem program, int index)
        {
            var row = new Panel
            {
                Height = 50,
                Dock = DockStyle.Top,
                BackColor = currentTheme == "dark" ? Color.FromArgb(45, 45, 45) : Color.White,
                Cursor = Cursors.Hand,
                Tag = index
            };

            // Bottom border
            var border = new Panel
            {
                Height = 1,
                Dock = DockStyle.Bottom,
                BackColor = currentTheme == "dark" ? Color.FromArgb(60, 60, 60) : Color.FromArgb(230, 230, 230)
            };

            // Icon
            var iconBox = new PictureBox
            {
                Size = new Size(32, 32),
                Location = new Point(8, 9),
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Transparent
            };
            try
            {
                if (program.Icon != null)
                    iconBox.Image = program.Icon.ToBitmap();
            }
            catch { }

            // Program name
            var nameLabel = new Label
            {
                Text = program.Name,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = currentTheme == "dark" ? Color.White : Color.FromArgb(30, 30, 30),
                Location = new Point(48, 4),
                AutoSize = false,
                Size = new Size(320, 22),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // File path
            var pathLabel = new Label
            {
                Text = Path.GetFileName(program.FilePath),
                Font = new Font("Segoe UI", 8),
                ForeColor = currentTheme == "dark" ? Color.FromArgb(180, 180, 180) : Color.Gray,
                Location = new Point(48, 26),
                AutoSize = false,
                Size = new Size(320, 18),
                TextAlign = ContentAlignment.MiddleLeft
            };

            row.Controls.Add(iconBox);
            row.Controls.Add(nameLabel);
            row.Controls.Add(pathLabel);
            row.Controls.Add(border);

            // Click events for selection
            Action<object?, MouseEventArgs> clickHandler = (s, args) =>
            {
                int idx = (int)row.Tag;
                if (Control.ModifierKeys.HasFlag(Keys.Control))
                {
                    // Toggle selection
                    if (selectedIndices.Contains(idx))
                    {
                        selectedIndices.Remove(idx);
                        row.BackColor = Color.White;
                        nameLabel.BackColor = Color.White;
                        pathLabel.BackColor = Color.White;
                    }
                    else
                    {
                        selectedIndices.Add(idx);
                        row.BackColor = Color.FromArgb(0, 120, 215);
                        nameLabel.BackColor = Color.FromArgb(0, 120, 215);
                        nameLabel.ForeColor = Color.White;
                        pathLabel.BackColor = Color.FromArgb(0, 120, 215);
                        pathLabel.ForeColor = Color.FromArgb(200, 200, 255);
                    }
                }
                else
                {
                    // Single selection - deselect all, select this one
                    DeselectAllRows();
                    selectedIndices.Clear();
                    selectedIndices.Add(idx);
                    row.BackColor = Color.FromArgb(0, 120, 215);
                    nameLabel.BackColor = Color.FromArgb(0, 120, 215);
                    nameLabel.ForeColor = Color.White;
                    pathLabel.BackColor = Color.FromArgb(0, 120, 215);
                    pathLabel.ForeColor = Color.FromArgb(200, 200, 255);
                }

                // Double-click: start program
                if (args is MouseEventArgs me && me.Clicks >= 2)
                {
                    var filtered = currentCategory == "Alla program"
                        ? programs.ToList()
                        : programs.Where(p => p.Category == currentCategory).ToList();
                    if (idx < filtered.Count)
                        filtered[idx].Start();
                }
            };

            row.MouseClick += (s, e) => clickHandler(s, e);
            iconBox.MouseClick += (s, e) => clickHandler(s, e);
            nameLabel.MouseClick += (s, e) => clickHandler(s, e);
            pathLabel.MouseClick += (s, e) => clickHandler(s, e);

            row.MouseDoubleClick += (s, e) => clickHandler(s, new MouseEventArgs(e.Button, 2, e.X, e.Y, e.Delta));
            nameLabel.MouseDoubleClick += (s, e) => clickHandler(s, new MouseEventArgs(e.Button, 2, e.X, e.Y, e.Delta));
            pathLabel.MouseDoubleClick += (s, e) => clickHandler(s, new MouseEventArgs(e.Button, 2, e.X, e.Y, e.Delta));
            iconBox.MouseDoubleClick += (s, e) => clickHandler(s, new MouseEventArgs(e.Button, 2, e.X, e.Y, e.Delta));

            // Hover effect
            Action<Control> addHover = null!;
            addHover = (ctrl) =>
            {
                ctrl.MouseEnter += (s, e) =>
                {
                    if (!selectedIndices.Contains((int)row.Tag))
                        row.BackColor = Color.FromArgb(235, 240, 255);
                };
                ctrl.MouseLeave += (s, e) =>
                {
                    if (!selectedIndices.Contains((int)row.Tag))
                        row.BackColor = currentTheme == "dark" ? Color.FromArgb(45, 45, 45) : Color.White;
                };
            };
            addHover(row);
            addHover(iconBox);
            addHover(nameLabel);
            addHover(pathLabel);

            // Drag & Drop
            row.AllowDrop = true;
            
            MouseEventHandler mouseDown = (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    draggedItem = program;
                    dragStartPoint = e.Location;
                }
            };

            MouseEventHandler mouseMove = (s, e) =>
            {
                if (e.Button == MouseButtons.Left && draggedItem != null)
                {
                    if (Math.Abs(e.X - dragStartPoint.X) > 5 || Math.Abs(e.Y - dragStartPoint.Y) > 5)
                    {
                        row.DoDragDrop(draggedItem, DragDropEffects.Move);
                    }
                }
            };

            DragEventHandler dragEnter = (s, e) =>
            {
                if (e.Data != null && e.Data.GetDataPresent(typeof(ProgramItem)))
                {
                    e.Effect = DragDropEffects.Move;
                }
            };

            DragEventHandler dragDrop = (s, e) =>
            {
                if (e.Data == null) return;
                var source = e.Data.GetData(typeof(ProgramItem)) as ProgramItem;
                if (source != null && source != program)
                {
                    MoveProgram(source, program);
                    draggedItem = null;
                }
            };

            // Attach drag events
            Control[] dragControls = { row, iconBox, nameLabel, pathLabel };
            foreach (var c in dragControls)
            {
                c.MouseDown += mouseDown;
                c.MouseMove += mouseMove;
            }
            
            row.DragEnter += dragEnter;
            row.DragDrop += dragDrop;

            // Right-click context menu
            var contextMenu = new ContextMenuStrip();

            // "Move to category" menu item with submenu
            var moveToItem = new ToolStripMenuItem(T("moveToCategory"));
            
            // Add all available categories (except the current one)
            foreach (var category in categories)
            {
                if (category.Name != program.Category)
                {
                    string displayName = category.Name == DefaultCategory ? T("allPrograms") : category.Name;
                    var categoryItem = new ToolStripMenuItem(displayName);
                    categoryItem.Click += (s, e) => 
                    {
                        program.Category = category.Name;
                        SavePrograms();
                        RefreshProgramList();
                    };
                    moveToItem.DropDownItems.Add(categoryItem);
                }
            }

            contextMenu.Items.Add(moveToItem);
            
            // "Rename" menu item
            var renameItem = new ToolStripMenuItem(T("contextRename"));
            renameItem.Click += (s, e) => ShowRenameDialog(program);
            contextMenu.Items.Add(renameItem);

            // "Delete" menu item
            var deleteItem = new ToolStripMenuItem(T("contextDelete"));
            deleteItem.Click += (s, e) => 
            {
                programs.Remove(program);
                SavePrograms();
                RefreshProgramList();
            };
            contextMenu.Items.Add(deleteItem);

            // Attach context menu to all controls in the row
            row.ContextMenuStrip = contextMenu;
            iconBox.ContextMenuStrip = contextMenu;
            nameLabel.ContextMenuStrip = contextMenu;
            pathLabel.ContextMenuStrip = contextMenu;

            return row;
        }

        private void DeselectAllRows()
        {
            if (programPanel == null) return;
            Color panelBg = currentTheme == "dark" ? Color.FromArgb(45, 45, 45) : Color.White;
            Color textColor = currentTheme == "dark" ? Color.White : Color.FromArgb(30, 30, 30);
            Color secondaryTextColor = currentTheme == "dark" ? Color.FromArgb(180, 180, 180) : Color.Gray;

            foreach (Control ctrl in programPanel.Controls)
            {
                if (ctrl is Panel p && p.Tag is int)
                {
                    p.BackColor = panelBg;
                    foreach (Control child in p.Controls)
                    {
                        if (child is Label lbl)
                        {
                            lbl.BackColor = panelBg;
                            if (lbl.Font.Bold)
                                lbl.ForeColor = textColor;
                            else
                                lbl.ForeColor = secondaryTextColor;
                        }
                    }
                }
            }
        }

        private void UpdateStatus()
        {
            if (statusLabel != null)
                statusLabel.Text = string.Format(T("programCount"), programs.Count, MaxPrograms);
        }

        private void ShowRenameDialog(ProgramItem program)
        {
            using (var inputDialog = new Form())
            {
                inputDialog.Text = T("editNameTitle");
                inputDialog.Size = new Size(350, 150);
                inputDialog.StartPosition = FormStartPosition.CenterParent;
                inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputDialog.MaximizeBox = false;
                inputDialog.MinimizeBox = false;

                var label = new Label { Text = T("newName"), Left = 20, Top = 20, Width = 100 };
                var textBox = new TextBox { Text = program.Name, Left = 120, Top = 20, Width = 200 };
                var okButton = new Button { Text = T("ok"), Left = 120, Top = 60, Width = 80, DialogResult = DialogResult.OK };
                var cancelButton = new Button { Text = T("cancel"), Left = 210, Top = 60, Width = 80, DialogResult = DialogResult.Cancel };

                inputDialog.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
                inputDialog.AcceptButton = okButton;
                inputDialog.CancelButton = cancelButton;

                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    string newName = textBox.Text.Trim();
                    if (!string.IsNullOrEmpty(newName))
                    {
                        program.Name = newName;
                        RefreshProgramList();
                        SavePrograms();
                    }
                }
            }
        }

        // ========== Button Event Handlers ==========

        private void AddButton_Click(object? sender, EventArgs e)
        {
            if (programs.Count >= MaxPrograms)
            {
                MessageBox.Show(string.Format(T("limitReached"), MaxPrograms), T("title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = T("selectFile");
                openFileDialog.Filter = T("fileFilter");
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        if (programs.Count >= MaxPrograms) break;

                        if (!programs.Any(p => p.FilePath.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var program = new ProgramItem(Path.GetFileNameWithoutExtension(fileName), fileName);
                            program.Category = currentCategory;
                            programs.Insert(0, program);
                        }
                    }

                    RefreshProgramList();
                    SavePrograms();
                }
            }
        }

        private void EditButton_Click(object? sender, EventArgs e)
        {
            if (selectedIndices.Count == 0)
            {
                MessageBox.Show(T("selectProgramEdit"), T("noSelection"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int idx = selectedIndices.First();
            var filtered = currentCategory == "Alla program"
                ? programs.ToList()
                : programs.Where(p => p.Category == currentCategory).ToList();
            if (idx >= filtered.Count) return;

            var program = filtered[idx];

            using (var inputDialog = new Form())
            {
                inputDialog.Text = T("editNameTitle");
                inputDialog.Size = new Size(350, 150);
                inputDialog.StartPosition = FormStartPosition.CenterParent;
                inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputDialog.MaximizeBox = false;
                inputDialog.MinimizeBox = false;

                var label = new Label { Text = T("newName"), Left = 20, Top = 20, Width = 100 };
                var textBox = new TextBox { Text = program.Name, Left = 120, Top = 20, Width = 200 };
                var okButton = new Button { Text = T("ok"), Left = 120, Top = 60, Width = 80, DialogResult = DialogResult.OK };
                var cancelButton = new Button { Text = T("cancel"), Left = 210, Top = 60, Width = 80, DialogResult = DialogResult.Cancel };

                inputDialog.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
                inputDialog.AcceptButton = okButton;
                inputDialog.CancelButton = cancelButton;

                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    string newName = textBox.Text.Trim();
                    if (!string.IsNullOrEmpty(newName))
                    {
                        program.Name = newName;
                        RefreshProgramList();
                        SavePrograms();
                    }
                }
            }
        }

        private void CategoryButton_Click(object? sender, EventArgs e)
        {
            using (var inputDialog = new Form())
            {
                inputDialog.Text = T("addCategoryTitle");
                inputDialog.Size = new Size(350, 150);
                inputDialog.StartPosition = FormStartPosition.CenterParent;
                inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputDialog.MaximizeBox = false;
                inputDialog.MinimizeBox = false;

                var label = new Label { Text = T("categoryName"), Left = 20, Top = 20, Width = 100 };
                var textBox = new TextBox { Left = 120, Top = 20, Width = 200 };
                var okButton = new Button { Text = T("ok"), Left = 120, Top = 60, Width = 80, DialogResult = DialogResult.OK };
                var cancelButton = new Button { Text = T("cancel"), Left = 210, Top = 60, Width = 80, DialogResult = DialogResult.Cancel };

                inputDialog.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
                inputDialog.AcceptButton = okButton;
                inputDialog.CancelButton = cancelButton;

                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    string categoryName = textBox.Text.Trim();
                    if (!string.IsNullOrEmpty(categoryName) && !categories.Any(c => c.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase)))
                    {
                        categories.Add(new Category(categoryName));
                        SaveCategories();
                        LoadCategories();
                        MessageBox.Show(string.Format(T("categoryAdded"), categoryName), T("title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (!string.IsNullOrEmpty(categoryName))
                    {
                        MessageBox.Show(T("categoryExists"), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void RemoveButton_Click(object? sender, EventArgs e)
        {
            if (selectedIndices.Count == 0)
            {
                MessageBox.Show(T("selectProgramRemove"), T("noSelection"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var filtered = currentCategory == "Alla program"
                ? programs.ToList()
                : programs.Where(p => p.Category == currentCategory).ToList();

            var toRemove = selectedIndices.OrderByDescending(i => i).Where(i => i < filtered.Count).Select(i => filtered[i]).ToList();
            foreach (var p in toRemove)
                programs.Remove(p);

            RefreshProgramList();
            SavePrograms();
        }

        private void RemoveCategoryButton_Click(object? sender, EventArgs e)
        {
            if (categoryComboBox == null || categoryComboBox.SelectedItem == null)
            {
                MessageBox.Show(T("selectCategoryRemove"), T("noSelection"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string selectedCategory = categoryComboBox.SelectedItem.ToString() ?? string.Empty;
            string displayDefault = T("allPrograms");

            if (selectedCategory == displayDefault || selectedCategory == DefaultCategory)
            {
                MessageBox.Show(T("cannotRemoveDefault"), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var programsInCategory = programs.Where(p => p.Category == selectedCategory).ToList();
            if (programsInCategory.Any())
            {
                var result = MessageBox.Show(
                    string.Format(T("programsInCategory"), programsInCategory.Count, selectedCategory),
                    T("title"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (var program in programsInCategory)
                        program.Category = DefaultCategory;
                    SavePrograms();
                }
                else return;
            }

            var categoryToRemove = categories.FirstOrDefault(c => c.Name == selectedCategory);
            if (categoryToRemove != null)
            {
                categories.Remove(categoryToRemove);
                SaveCategories();
                LoadCategories();
                MessageBox.Show(string.Format(T("categoryRemoved"), selectedCategory), T("title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void EditCategoryButton_Click(object? sender, EventArgs e)
        {
            if (categoryComboBox == null || categoryComboBox.SelectedItem == null)
            {
                MessageBox.Show(T("selectCategoryEdit"), T("noSelection"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string selectedCategory = categoryComboBox.SelectedItem.ToString() ?? string.Empty;
            string displayDefault = T("allPrograms");

            if (selectedCategory == displayDefault || selectedCategory == DefaultCategory)
            {
                MessageBox.Show(T("cannotEditDefault"), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var inputDialog = new Form())
            {
                inputDialog.Text = T("editCategoryTitle");
                inputDialog.Size = new Size(350, 150);
                inputDialog.StartPosition = FormStartPosition.CenterParent;
                inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputDialog.MaximizeBox = false;
                inputDialog.MinimizeBox = false;

                var label = new Label { Text = T("categoryName"), Left = 20, Top = 20, Width = 100 };
                var textBox = new TextBox { Left = 120, Top = 20, Width = 200, Text = selectedCategory };
                var okButton = new Button { Text = T("ok"), Left = 120, Top = 60, Width = 80, DialogResult = DialogResult.OK };
                var cancelButton = new Button { Text = T("cancel"), Left = 210, Top = 60, Width = 80, DialogResult = DialogResult.Cancel };

                inputDialog.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
                inputDialog.AcceptButton = okButton;
                inputDialog.CancelButton = cancelButton;

                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    string newCategoryName = textBox.Text.Trim();
                    if (string.IsNullOrEmpty(newCategoryName))
                        return;

                    if (newCategoryName.Equals(selectedCategory, StringComparison.OrdinalIgnoreCase))
                        return;

                    if (categories.Any(c => c.Name.Equals(newCategoryName, StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show(T("categoryExists"), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Update category name
                    var categoryToEdit = categories.FirstOrDefault(c => c.Name == selectedCategory);
                    if (categoryToEdit != null)
                    {
                        categoryToEdit.Name = newCategoryName;
                        
                        // Update programs
                        foreach (var prog in programs.Where(p => p.Category == selectedCategory))
                        {
                            prog.Category = newCategoryName;
                        }

                        SaveCategories();
                        SavePrograms();
                        
                        // Reload and select the new name
                        LoadCategories();
                        
                        if (categoryComboBox != null)
                        {
                            int index = -1;
                            for(int i=0; i < categoryComboBox.Items.Count; i++)
                            {
                                if (categoryComboBox.Items[i]?.ToString() == newCategoryName)
                                {
                                    index = i;
                                    break;
                                }
                            }
                            if (index >= 0)
                                categoryComboBox.SelectedIndex = index;
                        }

                        MessageBox.Show(string.Format(T("categoryRenamed"), selectedCategory, newCategoryName), T("title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void CategoryComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (categoryComboBox != null && categoryComboBox.SelectedIndex >= 0)
            {
                string selected = categoryComboBox.SelectedItem?.ToString() ?? DefaultCategory;
                if (selected == T("allPrograms"))
                    currentCategory = DefaultCategory;
                else
                    currentCategory = selected;
                    
                RefreshProgramList();
            }
        }

        // ========== Data persistence ==========

        private void LoadCategories()
        {
            try
            {
                var loadedCategories = new List<Category>();
                if (File.Exists(CategoryFile))
                {
                    var json = File.ReadAllText(CategoryFile);
                    var saved = JsonSerializer.Deserialize<List<Category>>(json);
                    if (saved != null)
                    {
                        loadedCategories.AddRange(saved);
                    }
                }

                // Filter out any existing "All programs" variants from loaded categories
                // Also filter "Alla program" to avoid confusion/duplication
                categories = loadedCategories.Where(c => 
                    !string.Equals(c.Name, DefaultCategory, StringComparison.OrdinalIgnoreCase) && 
                    !string.Equals(c.Name, "All programs", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.Name, "Alla program", StringComparison.OrdinalIgnoreCase)).ToList();

                // Always insert DefaultCategory at the top
                categories.Insert(0, new Category(DefaultCategory));

                if (categoryComboBox != null)
                {
                    categoryComboBox.SelectedIndexChanged -= CategoryComboBox_SelectedIndexChanged;
                    categoryComboBox.Items.Clear();
                    
                    // Use a HashSet to track added names to prevent UI duplicates
                    var addedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var category in categories)
                    {
                        string displayName = category.Name == DefaultCategory ? T("allPrograms") : category.Name;
                        
                        // Only add if not already present (case-insensitive check for display name)
                        if (!addedNames.Contains(displayName))
                        {
                            categoryComboBox.Items.Add(displayName);
                            addedNames.Add(displayName);
                        }
                    }

                    if (categoryComboBox.Items.Count > 0)
                        categoryComboBox.SelectedIndex = 0;
                        
                    currentCategory = DefaultCategory;
                    categoryComboBox.SelectedIndexChanged += CategoryComboBox_SelectedIndexChanged;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(T("couldNotLoad"), ex.Message), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                categories.Clear();
                categories.Add(new Category("All programs"));

                if (categoryComboBox != null)
                {
                    categoryComboBox.Items.Clear();
                    categoryComboBox.Items.Add("All programs");
                    categoryComboBox.SelectedItem = "All programs";
                    currentCategory = "All programs";
                }
            }
        }

        private void SaveCategories()
        {
            try
            {
                var json = JsonSerializer.Serialize(categories, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(CategoryFile, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(T("couldNotSave"), ex.Message), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LoadPrograms()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    var json = File.ReadAllText(ConfigFile);
                    var savedPrograms = JsonSerializer.Deserialize<List<ProgramData>>(json);
                    if (savedPrograms != null)
                    {
                        programs.Clear();
                        foreach (var data in savedPrograms)
                        {
                            var program = new ProgramItem(data.Name, data.FilePath, data.Arguments);
                            program.AddedDate = data.AddedDate;
                            program.Category = data.Category ?? DefaultCategory;
                            if (string.IsNullOrEmpty(program.Category) || program.Category == "Alla program")
                                program.Category = DefaultCategory;
                            programs.Add(program);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(T("couldNotLoad"), ex.Message), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            RefreshProgramList();
        }

        private void SavePrograms()
        {
            try
            {
                var programData = programs.Select(p => new ProgramData
                {
                    Name = p.Name,
                    FilePath = p.FilePath,
                    Arguments = p.Arguments,
                    AddedDate = p.AddedDate,
                    Category = p.Category
                }).ToList();

                var json = JsonSerializer.Serialize(programData, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigFile, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(T("couldNotSave"), ex.Message), T("error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ========== Window behavior ==========

        protected override void OnResize(EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                this.Hide();
            }
            base.OnResize(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
            else
            {
                notifyIcon?.Dispose();
            }
            base.OnFormClosing(e);
        }
    }

    public class ProgramData
    {
        public string Name { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string? Arguments { get; set; }
        public DateTime AddedDate { get; set; }
        public string? Category { get; set; }
    }

    public class AppSettings
    {
        public string? Language { get; set; }
        public string? Theme { get; set; }
        public bool? Autostart { get; set; }
    }
}

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace SharpMZBasicProgramEditor
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            var app = new Application();
            app.Run(new MainWindow());
        }
    }

    internal sealed class MainWindow : Window
    {
        private TextBox _editor;
        private TextBox _programNameText;
        private TextBlock _statusText;
        private ComboBox _modeCombo;
        private FontFamily _fontFamily;

        private string _currentFilePath;
        private MzfHeader _loadedHeader;

        public MainWindow()
        {
            Title = "Sharp MZ Basic Program Editor";
            Width = 1180;
            Height = 760;
            MinWidth = 920;
            MinHeight = 620;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(Color.FromRgb(0xE7, 0xDE, 0xD2));
            AllowDrop = true;

            string fontPath;
            if (!TryExtractEmbeddedFont(out fontPath))
            {
                MessageBox.Show("Embedded SharpMZ font could not be loaded.", "Font Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                Close();
                return;
            }

            _fontFamily = new FontFamily(new Uri(Path.GetDirectoryName(fontPath) + Path.DirectorySeparatorChar, UriKind.Absolute), "./#SharpMZ");

            var root = new Grid { Margin = new Thickness(18) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var toolbar = BuildToolbar();
            Grid.SetRow(toolbar, 0);
            root.Children.Add(toolbar);

            var content = BuildMainContent();
            Grid.SetRow(content, 1);
            root.Children.Add(content);

            _statusText = new TextBlock
            {
                Margin = new Thickness(2, 12, 2, 0),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(0x32, 0x38, 0x42)),
                Text = "Open a BASIC .mzf file, or start typing a new program."
            };
            Grid.SetRow(_statusText, 2);
            root.Children.Add(_statusText);

            Content = root;

            DragOver += OnDragOver;
            Drop += OnDrop;

            NewProgram();
        }

        private static bool TryExtractEmbeddedFont(out string fontPath)
        {
            fontPath = null;

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("SharpMZ.ttf"))
                {
                    if (stream == null)
                    {
                        return false;
                    }

                    var cacheDir = Path.Combine(Path.GetTempPath(), "SharpMZBasicProgramEditor");
                    Directory.CreateDirectory(cacheDir);

                    fontPath = Path.Combine(cacheDir, "SharpMZ.ttf");
                    using (var file = new FileStream(fontPath, FileMode.Create, FileAccess.Write, FileShare.Read))
                    {
                        stream.CopyTo(file);
                    }

                    return true;
                }
            }
            catch
            {
                fontPath = null;
                return false;
            }
        }

        private UIElement BuildToolbar()
        {
            var panel = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x24, 0x2B, 0x35)),
                CornerRadius = new CornerRadius(18),
                Padding = new Thickness(18)
            };

            var grid = new Grid();
            for (var i = 0; i < 7; i++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(12) });
            }
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var newButton = CreateToolbarButton("New");
            newButton.Click += delegate { NewProgram(); };
            Grid.SetColumn(newButton, 0);
            grid.Children.Add(newButton);

            var openButton = CreateToolbarButton("Open .MZF");
            openButton.Click += OnOpenFileClicked;
            Grid.SetColumn(openButton, 2);
            grid.Children.Add(openButton);

            var saveButton = CreateToolbarButton("Save .MZF");
            saveButton.Click += OnSaveFileClicked;
            Grid.SetColumn(saveButton, 4);
            grid.Children.Add(saveButton);

            var checkButton = CreateToolbarButton("Check S-BASIC");
            checkButton.Click += OnCheckSyntaxClicked;
            Grid.SetColumn(checkButton, 6);
            grid.Children.Add(checkButton);

            var modeLabel = CreateToolbarLabel("Dialect");
            Grid.SetColumn(modeLabel, 8);
            grid.Children.Add(modeLabel);

            _modeCombo = new ComboBox
            {
                Width = 150,
                FontSize = 14,
                SelectedIndex = 0,
                ItemsSource = new object[]
                {
                    BasicDialect.SBasic,
                    BasicDialect.HuBasic
                }
            };
            Grid.SetColumn(_modeCombo, 10);
            grid.Children.Add(_modeCombo);

            var nameLabel = CreateToolbarLabel("Program Name");
            Grid.SetColumn(nameLabel, 12);
            grid.Children.Add(nameLabel);

            _programNameText = new TextBox
            {
                Width = 180,
                Margin = new Thickness(-6, 0, 0, 0),
                FontSize = 14,
                VerticalContentAlignment = VerticalAlignment.Center,
                Text = "NEWPROGRAM"
            };
            Grid.SetColumn(_programNameText, 14);
            grid.Children.Add(_programNameText);

            panel.Child = grid;
            return panel;
        }

        private UIElement BuildMainContent()
        {
            var grid = new Grid { Margin = new Thickness(0, 16, 0, 0) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });

            _editor = new TextBox
            {
                AcceptsReturn = true,
                AcceptsTab = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                TextWrapping = TextWrapping.NoWrap,
                FontSize = 20,
                FontFamily = _fontFamily,
                Background = new SolidColorBrush(Color.FromRgb(0x12, 0x18, 0x1E)),
                Foreground = new SolidColorBrush(Color.FromRgb(0xF2, 0xEE, 0xDC)),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(18)
            };
            _editor.PreviewKeyDown += OnEditorPreviewKeyDown;
            _editor.PreviewTextInput += OnEditorPreviewTextInput;
            Grid.SetColumn(_editor, 0);
            grid.Children.Add(_editor);

            var palette = BuildPalettePanel();
            Grid.SetColumn(palette, 2);
            grid.Children.Add(palette);

            return grid;
        }

        private UIElement BuildPalettePanel()
        {
            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x24, 0x2B, 0x35)),
                CornerRadius = new CornerRadius(18),
                Padding = new Thickness(14)
            };

            var dock = new DockPanel();

            var info = new TextBlock
            {
                Text = "Character palette. Click any cell to insert that Sharp MZ character at the caret.",
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(0xD4, 0xDD, 0xE6)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            };
            DockPanel.SetDock(info, Dock.Top);
            dock.Children.Add(info);

            var tabs = new TabControl();
            tabs.Items.Add(CreateCharacterTab("Charset 1", 0));
            tabs.Items.Add(CreateCharacterTab("Charset 2", 1));
            tabs.Items.Add(CreateCharacterTab("Charset 3", 2));
            dock.Children.Add(tabs);

            border.Child = dock;
            return border;
        }

        private TabItem CreateCharacterTab(string title, int tableIndex)
        {
            var grid = new UniformGrid
            {
                Columns = 16,
                Rows = 16
            };

            for (var code = 0; code < 256; code++)
            {
                var button = new Button
                {
                    Margin = new Thickness(1),
                    Padding = new Thickness(0),
                    Height = 30,
                    FontSize = 16,
                    FontFamily = _fontFamily,
                    Background = new SolidColorBrush(Color.FromRgb(0x14, 0x1A, 0x22)),
                    Foreground = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0x41, 0x4D, 0x5A)),
                    Content = SharpMzTextEncoding.GetPaletteCharacter(code, tableIndex).ToString()
                };

                var codeCopy = code;
                var tableCopy = tableIndex;
                button.ToolTip = string.Format(CultureInfo.InvariantCulture, "0x{0:X2}  table {1}", codeCopy, tableCopy + 1);
                button.Click += delegate { InsertPaletteCharacter(codeCopy, tableCopy); };
                grid.Children.Add(button);
            }

            return new TabItem
            {
                Header = title,
                Content = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Content = grid
                }
            };
        }

        private static Button CreateToolbarButton(string title)
        {
            return new Button
            {
                Content = title,
                Padding = new Thickness(16, 10, 16, 10),
                FontSize = 15,
                Background = new SolidColorBrush(Color.FromRgb(0xCF, 0x84, 0x44)),
                Foreground = Brushes.White,
                BorderBrush = Brushes.Transparent
            };
        }

        private static TextBlock CreateToolbarLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(Color.FromRgb(0xD6, 0xDE, 0xE8)),
                FontSize = 14
            };
        }

        private void OnOpenFileClicked(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Open Sharp MZ BASIC Program",
                Filter = "Sharp MZ BASIC (*.mzf;*.mzt)|*.mzf;*.mzt|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) == true)
            {
                LoadProgram(dialog.FileName);
            }
        }

        private void OnSaveFileClicked(object sender, RoutedEventArgs e)
        {
            var defaultName = _programNameText.Text.Trim();
            if (defaultName.Length == 0)
            {
                defaultName = "NEWPROGRAM";
            }

            var dialog = new SaveFileDialog
            {
                Title = "Save Sharp MZ BASIC Program",
                Filter = "Sharp MZ BASIC (*.mzf)|*.mzf|Sharp MZ tape image (*.mzt)|*.mzt|All files (*.*)|*.*",
                FileName = defaultName + ".mzf"
            };

            if (!string.IsNullOrEmpty(_currentFilePath))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(_currentFilePath);
            }

            if (dialog.ShowDialog(this) == true)
            {
                SaveProgram(dialog.FileName);
            }
        }

        private void OnCheckSyntaxClicked(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialect = (BasicDialect)_modeCombo.SelectedItem;
                var issues = BasicSyntaxChecker.CheckEditorText(_editor.Text, dialect);
                if (issues.Count == 0)
                {
                    UpdateStatus("No syntax issues found.");
                    MessageBox.Show(this, "No syntax issues found.", "S-BASIC Check", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var builder = new StringBuilder();
                var limit = Math.Min(issues.Count, 20);
                for (var i = 0; i < limit; i++)
                {
                    builder.Append("Line ");
                    builder.Append(issues[i].LineNumber.ToString(CultureInfo.InvariantCulture));
                    builder.Append(": ");
                    builder.AppendLine(issues[i].Message);
                }

                if (issues.Count > limit)
                {
                    builder.Append("...and ");
                    builder.Append((issues.Count - limit).ToString(CultureInfo.InvariantCulture));
                    builder.Append(" more.");
                }

                UpdateStatus(string.Format(CultureInfo.InvariantCulture, "Syntax checker found {0} issue(s).", issues.Count));
                MessageBox.Show(this, builder.ToString(), "S-BASIC Check", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                UpdateStatus(ex.Message);
                MessageBox.Show(this, ex.Message, "S-BASIC Check Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NewProgram()
        {
            _currentFilePath = null;
            _loadedHeader = null;
            _programNameText.Text = "NEWPROGRAM";
            _modeCombo.SelectedItem = BasicDialect.SBasic;
            _editor.Text = SharpMzTextEncoding.ToDisplayText("10 REM NEW PROGRAM");
            _editor.CaretIndex = _editor.Text.Length;
            UpdateStatus("New program ready.");
        }

        private void OnDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void OnDrop(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files != null && files.Length > 0)
            {
                LoadProgram(files[0]);
            }
        }

        private void LoadProgram(string filePath)
        {
            try
            {
                var bytes = File.ReadAllBytes(filePath);
                var header = MzfFile.ReadHeader(bytes);
                var requestedDialect = (BasicDialect)_modeCombo.SelectedItem;
                var dialect = requestedDialect;
                var program = TryDecodeWithFallback(bytes, requestedDialect, out dialect);
                _currentFilePath = filePath;
                _loadedHeader = header;

                var builder = new StringBuilder();
                for (var i = 0; i < program.Count; i++)
                {
                    if (builder.Length > 0)
                    {
                        builder.Append(Environment.NewLine);
                    }

                    builder.Append(SharpMzTextEncoding.ToDisplayText(program[i].Number.ToString(CultureInfo.InvariantCulture)));
                    if (program[i].DisplayBytes.Length > 0)
                    {
                        builder.Append(SharpMzTextEncoding.ToDisplayCharacter(' '));
                        builder.Append(SharpMzTextEncoding.FromSharpBytes(program[i].DisplayBytes));
                    }
                }

                _programNameText.Text = header.Name;
                _editor.Text = builder.ToString();
                _modeCombo.SelectedItem = dialect;

                UpdateStatus(string.Format(
                    CultureInfo.InvariantCulture,
                    requestedDialect == dialect ? "{0} loaded. {1} lines, mode {2}." : "{0} loaded. {1} lines, auto-switched to {2}.",
                    Path.GetFileName(filePath),
                    program.Count,
                    dialect == BasicDialect.SBasic ? "S-BASIC" : "HuBASIC"));
            }
            catch (Exception ex)
            {
                UpdateStatus(ex.Message);
                MessageBox.Show(ex.Message, "Load Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<BasicLine> TryDecodeWithFallback(byte[] bytes, BasicDialect requestedDialect, out BasicDialect resolvedDialect)
        {
            resolvedDialect = requestedDialect;

            try
            {
                return BasicDecoder.Decode(bytes, requestedDialect);
            }
            catch
            {
                var alternateDialect = requestedDialect == BasicDialect.SBasic ? BasicDialect.HuBasic : BasicDialect.SBasic;

                try
                {
                    var program = BasicDecoder.Decode(bytes, alternateDialect);
                    resolvedDialect = alternateDialect;
                    return program;
                }
                catch
                {
                    resolvedDialect = requestedDialect;
                    throw;
                }
            }
        }

        private void SaveProgram(string filePath)
        {
            try
            {
                var dialect = (BasicDialect)_modeCombo.SelectedItem;
                var lines = BasicEncoder.ParseEditorText(_editor.Text, dialect);
                var body = BasicEncoder.Encode(lines, dialect);
                var header = BuildHeaderForSave(dialect, body);

                File.WriteAllBytes(filePath, MzfFile.BuildFile(header, body));
                _currentFilePath = filePath;
                _loadedHeader = header;

                UpdateStatus(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0} saved. {1} lines, mode {2}.",
                    Path.GetFileName(filePath),
                    lines.Count,
                    dialect == BasicDialect.SBasic ? "S-BASIC" : "HuBASIC"));
            }
            catch (Exception ex)
            {
                UpdateStatus(ex.Message);
                MessageBox.Show(ex.Message, "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private MzfHeader BuildHeaderForSave(BasicDialect dialect, byte[] body)
        {
            var header = new MzfHeader();
            if (dialect == BasicDialect.HuBasic)
            {
                header.Type = _loadedHeader != null ? _loadedHeader.Type : (byte)2;
                header.LoadAddress = _loadedHeader != null ? _loadedHeader.LoadAddress : (ushort)0;
                header.ExecuteAddress = _loadedHeader != null ? _loadedHeader.ExecuteAddress : (ushort)0;
                header.NameTerminatedWithCarriageReturn = _loadedHeader != null && _loadedHeader.NameTerminatedWithCarriageReturn;
            }
            else
            {
                header.Type = 5;
                header.LoadAddress = 0x6BCF;
                header.ExecuteAddress = 0;
                header.NameTerminatedWithCarriageReturn = true;
            }

            header.Name = _programNameText.Text;
            header.Size = (ushort)body.Length;
            return header;
        }

        private void InsertPaletteCharacter(int code, int tableIndex)
        {
            var ch = SharpMzTextEncoding.GetPaletteCharacter(code, tableIndex).ToString();
            ReplaceEditorSelection(ch);
            _editor.Focus();
        }

        private void OnEditorPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Text))
            {
                return;
            }

            var shiftPressed = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            var builder = new StringBuilder();
            for (var i = 0; i < e.Text.Length; i++)
            {
                builder.Append(SharpMzTextEncoding.ToDisplayCharacter(e.Text[i], shiftPressed));
            }

            ReplaceEditorSelection(builder.ToString());
            e.Handled = true;
        }

        private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                ReplaceEditorSelection(SharpMzTextEncoding.ToDisplayCharacter(' ').ToString());
                e.Handled = true;
            }
        }

        private void ReplaceEditorSelection(string text)
        {
            var selectionStart = _editor.SelectionStart;
            var selectionLength = _editor.SelectionLength;
            var existingText = _editor.Text ?? string.Empty;

            if (selectionLength > 0)
            {
                existingText = existingText.Remove(selectionStart, selectionLength);
            }

            _editor.Text = existingText.Insert(selectionStart, text);
            _editor.SelectionStart = selectionStart + text.Length;
            _editor.SelectionLength = 0;
        }

        private void UpdateStatus(string text)
        {
            _statusText.Text = text;
        }
    }

    internal enum BasicDialect
    {
        SBasic,
        HuBasic
    }

    internal sealed class BasicLine
    {
        public int Number;
        public byte[] DisplayBytes;
    }

    internal sealed class MzfHeader
    {
        public byte Type;
        public string Name;
        public ushort Size;
        public ushort LoadAddress;
        public ushort ExecuteAddress;
        public bool NameTerminatedWithCarriageReturn;
    }

    internal static class MzfFile
    {
        public static MzfHeader ReadHeader(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 128)
            {
                throw new InvalidDataException("The file is too small to be a valid MZF file.");
            }

            var nameBuilder = new StringBuilder();
            var sawCarriageReturn = false;
            for (var i = 1; i <= 17; i++)
            {
                var value = bytes[i];
                if (value == 0x0D)
                {
                    sawCarriageReturn = true;
                    break;
                }

                if (value == 0)
                {
                    break;
                }

                if (value >= 32 && value <= 126)
                {
                    nameBuilder.Append((char)value);
                }
            }

            return new MzfHeader
            {
                Type = bytes[0],
                Name = nameBuilder.Length == 0 ? "PROGRAM" : nameBuilder.ToString(),
                Size = (ushort)(bytes[0x12] | (bytes[0x13] << 8)),
                LoadAddress = (ushort)(bytes[0x14] | (bytes[0x15] << 8)),
                ExecuteAddress = (ushort)(bytes[0x16] | (bytes[0x17] << 8)),
                NameTerminatedWithCarriageReturn = sawCarriageReturn
            };
        }

        public static byte[] BuildFile(MzfHeader header, byte[] body)
        {
            var file = new byte[128 + body.Length];
            file[0] = header.Type;

            for (var i = 0; i < 17; i++)
            {
                file[1 + i] = 0x20;
            }

            var name = (header.Name ?? "PROGRAM").ToUpperInvariant();
            if (name.Length > 16)
            {
                name = name.Substring(0, 16);
            }

            for (var i = 0; i < name.Length; i++)
            {
                file[1 + i] = (byte)name[i];
            }

            if (header.NameTerminatedWithCarriageReturn && name.Length < 17)
            {
                file[1 + name.Length] = 0x0D;
            }
            file[0x12] = (byte)(header.Size & 0xFF);
            file[0x13] = (byte)(header.Size >> 8);
            file[0x14] = (byte)(header.LoadAddress & 0xFF);
            file[0x15] = (byte)(header.LoadAddress >> 8);
            file[0x16] = (byte)(header.ExecuteAddress & 0xFF);
            file[0x17] = (byte)(header.ExecuteAddress >> 8);

            Buffer.BlockCopy(body, 0, file, 128, body.Length);
            return file;
        }
    }

    internal static class SharpMzTextEncoding
    {
        private const char DisplaySpace = '\u00A0';
        private static readonly Dictionary<char, byte> KeyboardToSharp = CreateKeyboardToSharpMap();
        private static readonly Dictionary<byte, char> SharpToKeyboard = CreateSharpToKeyboardMap();

        public static string FromSharpBytes(byte[] bytes)
        {
            var builder = new StringBuilder(bytes.Length);
            for (var i = 0; i < bytes.Length; i++)
            {
                builder.Append(ToEditorCharacter(bytes[i]));
            }

            return builder.ToString();
        }

        public static string ToDisplayText(string text)
        {
            var builder = new StringBuilder(text.Length);
            for (var i = 0; i < text.Length; i++)
            {
                builder.Append(ToDisplayCharacter(text[i]));
            }

            return builder.ToString();
        }

        public static char ToEditorCharacter(byte code)
        {
            if (code == 0x20)
            {
                return DisplaySpace;
            }

            return (char)(0xE000 + code);
        }

        public static char GetPaletteCharacter(int code, int tableIndex)
        {
            if (tableIndex == 0 && code == 0x20)
            {
                return DisplaySpace;
            }

            return (char)(0xE000 + (tableIndex * 0x100) + code);
        }

        public static byte ToSharpByte(char value)
        {
            return ToSharpByte(value, false);
        }

        public static byte ToSharpByte(char value, bool shiftLettersToCharset2)
        {
            byte mappedValue;
            if (KeyboardToSharp.TryGetValue(value, out mappedValue))
            {
                if (shiftLettersToCharset2 && value >= 'A' && value <= 'Z')
                {
                    return (byte)(0x81 + (value - 'A'));
                }

                return mappedValue;
            }

            if (value >= 'a' && value <= 'z')
            {
                return (byte)char.ToUpperInvariant(value);
            }

            if (shiftLettersToCharset2 && value >= 'A' && value <= 'Z')
            {
                return (byte)(0x81 + (value - 'A'));
            }

            if (value >= 32 && value <= 126)
            {
                return (byte)value;
            }

            if (value >= 0xE000 && value <= 0xE2FF)
            {
                return (byte)((int)value & 0xFF);
            }

            return (byte)'?';
        }

        public static char ToDisplayCharacter(char value)
        {
            return ToDisplayCharacter(value, false);
        }

        public static char ToDisplayCharacter(char value, bool shiftLettersToCharset2)
        {
            if (value == ' ')
            {
                return DisplaySpace;
            }

            if (shiftLettersToCharset2 && value >= 'A' && value <= 'Z')
            {
                return GetPaletteCharacter(0x81 + (value - 'A'), 1);
            }

            return ToEditorCharacter(ToSharpByte(value, shiftLettersToCharset2));
        }

        public static char ToLogicalCharacter(char value)
        {
            if (value == DisplaySpace)
            {
                return ' ';
            }

            if (value >= 0xE000 && value <= 0xE2FF)
            {
                var code = (byte)((int)value & 0xFF);
                char mapped;
                if (SharpToKeyboard.TryGetValue(code, out mapped))
                {
                    return mapped;
                }

                if (code >= 32 && code <= 126)
                {
                    return (char)code;
                }
            }

            return value;
        }

        public static string ToLogicalText(string text)
        {
            var builder = new StringBuilder(text.Length);
            for (var i = 0; i < text.Length; i++)
            {
                builder.Append(ToLogicalCharacter(text[i]));
            }

            return builder.ToString();
        }

        private static Dictionary<char, byte> CreateKeyboardToSharpMap()
        {
            var map = new Dictionary<char, byte>();

            for (var value = 0x30; value <= 0x39; value++)
            {
                map[(char)value] = (byte)value;
            }

            for (var value = 0x41; value <= 0x5A; value++)
            {
                map[(char)value] = (byte)value;
                map[char.ToLowerInvariant((char)value)] = (byte)value;
            }

            map['!'] = 0x21;
            map[' '] = 0x20;
            map['"'] = 0x22;
            map['“'] = 0x22;
            map['”'] = 0x22;
            map['#'] = 0x23;
            map['$'] = 0x24;
            map['%'] = 0x25;
            map['&'] = 0x26;
            map['('] = 0x28;
            map[')'] = 0x29;
            map['+'] = 0x2B;
            map[','] = 0x2C;
            map['-'] = 0x2D;
            map['–'] = 0x2D;
            map['.'] = 0x2E;
            map['/'] = 0x2F;
            map['*'] = 0x2A;
            map[':'] = 0x3A;
            map[';'] = 0x3B;
            map['<'] = 0x3C;
            map['>'] = 0x3E;
            map['?'] = 0x3F;
            map['@'] = 0x40;
            map['['] = 0x5B;
            map['\\'] = 0x5C;
            map[']'] = 0x5D;
            map['~'] = 0x94;
            map['|'] = 0xC0;

            return map;
        }

        private static Dictionary<byte, char> CreateSharpToKeyboardMap()
        {
            var map = new Dictionary<byte, char>();
            foreach (var pair in KeyboardToSharp)
            {
                if (!map.ContainsKey(pair.Value))
                {
                    map[pair.Value] = pair.Key;
                }
            }

            map[0x22] = '"';
            map[0x20] = ' ';
            map[0x2D] = '-';
            map[0x2A] = '*';
            map[0x3A] = ':';
            map[0x94] = '~';
            map[0xC0] = '|';

            return map;
        }
    }

    internal sealed class TokenInfo
    {
        public string Text;
        public byte[] Bytes;
    }

    internal sealed class SyntaxIssue
    {
        public int LineNumber;
        public string Message;
    }

    internal static class BasicSyntaxChecker
    {
        private static readonly HashSet<string> SBasicStatementHeads = CreateStatementHeads(BasicDecoder.TokenS0, BasicDecoder.TokenSFE);
        private static readonly HashSet<string> HuBasicStatementHeads = CreateStatementHeads(BasicDecoder.TokenH0, BasicDecoder.TokenHFE);
        private static readonly HashSet<string> SBasicFunctionNames = CreateFunctionNames(BasicDecoder.TokenSFF);
        private static readonly HashSet<string> HuBasicFunctionNames = CreateFunctionNames(BasicDecoder.TokenHFF);
        private static readonly string[] SBasicKeywordTokens = CreateKeywordTokens(SBasicStatementHeads, SBasicFunctionNames);
        private static readonly string[] HuBasicKeywordTokens = CreateKeywordTokens(HuBasicStatementHeads, HuBasicFunctionNames);

        private enum SyntaxTokenKind
        {
            Identifier,
            Keyword,
            Number,
            String,
            Operator,
            Comma,
            Semicolon,
            LeftParen,
            RightParen,
            LeftBracket,
            RightBracket,
            End
        }

        private sealed class SyntaxToken
        {
            public SyntaxTokenKind Kind;
            public string Text;
        }

        private sealed class ParserState
        {
            private readonly List<SyntaxToken> _tokens;
            private int _index;

            public ParserState(List<SyntaxToken> tokens)
            {
                _tokens = tokens;
                _index = 0;
            }

            public SyntaxToken Current
            {
                get { return _tokens[_index]; }
            }

            public SyntaxToken Previous
            {
                get { return _index > 0 ? _tokens[_index - 1] : _tokens[0]; }
            }

            public int Position
            {
                get { return _index; }
                set { _index = value < 0 ? 0 : (value >= _tokens.Count ? _tokens.Count - 1 : value); }
            }

            public bool Match(SyntaxTokenKind kind)
            {
                if (Current.Kind == kind)
                {
                    Consume();
                    return true;
                }

                return false;
            }

            public bool Match(SyntaxTokenKind kind, string text)
            {
                if (Current.Kind == kind && string.Equals(Current.Text, text, StringComparison.OrdinalIgnoreCase))
                {
                    Consume();
                    return true;
                }

                return false;
            }

            public SyntaxToken Consume()
            {
                var token = Current;
                if (_index < _tokens.Count - 1)
                {
                    _index++;
                }

                return token;
            }
        }

        public static List<SyntaxIssue> CheckEditorText(string editorText, BasicDialect dialect)
        {
            var issues = new List<SyntaxIssue>();
            var previousLineNumber = -1;
            var sourceLines = (editorText ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

            for (var i = 0; i < sourceLines.Length; i++)
            {
                var source = SharpMzTextEncoding.ToLogicalText(sourceLines[i]);
                if (string.IsNullOrWhiteSpace(source))
                {
                    continue;
                }

                var lineIssueCount = issues.Count;
                int lineNumber;
                string body;
                string parseError;
                if (!TryParseLine(source, out lineNumber, out body, out parseError))
                {
                    issues.Add(new SyntaxIssue { LineNumber = 0, Message = "Editor line " + (i + 1).ToString(CultureInfo.InvariantCulture) + ": " + parseError });
                    continue;
                }

                if (lineNumber <= previousLineNumber)
                {
                    issues.Add(new SyntaxIssue { LineNumber = lineNumber, Message = "Line numbers should be strictly increasing." });
                }
                previousLineNumber = lineNumber;

                CheckBody(lineNumber, body, issues, dialect);

                if (lineIssueCount == issues.Count && body.Length == 0)
                {
                    issues.Add(new SyntaxIssue { LineNumber = lineNumber, Message = "Line has no statement body." });
                }
            }

            return issues;
        }

        private static bool TryParseLine(string source, out int lineNumber, out string body, out string error)
        {
            lineNumber = 0;
            body = string.Empty;
            error = null;

            var trimmed = source.TrimStart();
            var cursor = 0;
            while (cursor < trimmed.Length && char.IsDigit(trimmed[cursor]))
            {
                cursor++;
            }

            if (cursor == 0)
            {
                error = "Missing BASIC line number.";
                return false;
            }

            if (!int.TryParse(trimmed.Substring(0, cursor), NumberStyles.Integer, CultureInfo.InvariantCulture, out lineNumber))
            {
                error = "Invalid BASIC line number.";
                return false;
            }

            if (lineNumber < 1 || lineNumber > 65535)
            {
                error = "Line number is out of range.";
                return false;
            }

            body = cursor < trimmed.Length ? trimmed.Substring(cursor).TrimStart() : string.Empty;
            return true;
        }

        private static void CheckBody(int lineNumber, string body, List<SyntaxIssue> issues, BasicDialect dialect)
        {
            if (body.Length == 0)
            {
                return;
            }

            var statements = SplitStatements(body);
            for (var i = 0; i < statements.Count; i++)
            {
                if (statements[i].Length == 0 && (i == 0 || i == statements.Count - 1))
                {
                    continue;
                }

                string error;
                if (!TryParseStatement(statements[i], dialect, out error))
                {
                    issues.Add(new SyntaxIssue { LineNumber = lineNumber, Message = error });
                }
            }
        }

        private static List<string> SplitStatements(string body)
        {
            var statements = new List<string>();
            var builder = new StringBuilder();
            var inQuotes = false;

            for (var i = 0; i < body.Length; i++)
            {
                var ch = body[i];
                if (ch == '"')
                {
                    inQuotes = !inQuotes;
                    builder.Append(ch);
                    continue;
                }

                if (!inQuotes && ch == ':')
                {
                    statements.Add(builder.ToString().Trim());
                    builder.Length = 0;
                    continue;
                }

                builder.Append(ch);
            }

            statements.Add(builder.ToString().Trim());
            return statements;
        }

        private static bool TryParseStatement(string statement, BasicDialect dialect, out string error)
        {
            error = null;
            if (statement.Length == 0)
            {
                return true;
            }

            if (statement.TrimStart().StartsWith("REM", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            List<SyntaxToken> tokens;
            if (!TryTokenize(statement, dialect, out tokens, out error))
            {
                return false;
            }

            var parser = new ParserState(tokens);
            if (!TryParseSingleStatement(parser, out error))
            {
                return false;
            }

            if (parser.Current.Kind != SyntaxTokenKind.End)
            {
                error = "Unexpected token '" + parser.Current.Text + "'.";
                return false;
            }

            return true;
        }

        private static bool TryTokenize(string statement, BasicDialect dialect, out List<SyntaxToken> tokens, out string error)
        {
            tokens = new List<SyntaxToken>();
            error = null;

            var i = 0;
            while (i < statement.Length)
            {
                var ch = statement[i];
                if (char.IsWhiteSpace(ch))
                {
                    i++;
                    continue;
                }

                if (ch < 32)
                {
                    error = "Unexpected control character in statement.";
                    return false;
                }

                if (ch == '"')
                {
                    var start = i++;
                    while (i < statement.Length && statement[i] != '"')
                    {
                        i++;
                    }

                    if (i >= statement.Length)
                    {
                        error = "Unterminated string literal.";
                        return false;
                    }

                    i++;
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.String, Text = statement.Substring(start, i - start) });
                    continue;
                }

                if (ch == ',')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.Comma, Text = "," });
                    i++;
                    continue;
                }

                if (ch == ';')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.Semicolon, Text = ";" });
                    i++;
                    continue;
                }

                if (ch == '(')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.LeftParen, Text = "(" });
                    i++;
                    continue;
                }

                if (ch == ')')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.RightParen, Text = ")" });
                    i++;
                    continue;
                }

                if (ch == '[')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.LeftBracket, Text = "[" });
                    i++;
                    continue;
                }

                if (ch == ']')
                {
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.RightBracket, Text = "]" });
                    i++;
                    continue;
                }

                if (IsOperatorStart(statement, i))
                {
                    var op = ReadOperator(statement, ref i);
                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.Operator, Text = op });
                    continue;
                }

                if (char.IsDigit(ch) || ch == '.' || ch == '$' || ch == '&')
                {
                    string numberText;
                    if (!TryReadNumberToken(statement, ref i, out numberText, out error))
                    {
                        return false;
                    }

                    tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.Number, Text = numberText });
                    continue;
                }

                if (IsIdentifierStart(ch))
                {
                    string keyword;
                    if (TryReadKeywordToken(statement, i, dialect, out keyword))
                    {
                        tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.Keyword, Text = keyword });
                        i += keyword.Length;
                        continue;
                    }

                    var start = i++;
                    while (i < statement.Length && IsIdentifierPart(statement[i]))
                    {
                        i++;
                    }

                    var text = statement.Substring(start, i - start);
                    var upper = text.ToUpperInvariant();
                    tokens.Add(new SyntaxToken
                    {
                        Kind = GetStatementHeads(dialect).Contains(upper) || GetFunctionNames(dialect).Contains(upper) ? SyntaxTokenKind.Keyword : SyntaxTokenKind.Identifier,
                        Text = upper
                    });
                    continue;
                }

                error = "Unexpected character '" + ch + "'.";
                return false;
            }

            tokens.Add(new SyntaxToken { Kind = SyntaxTokenKind.End, Text = string.Empty });
            return true;
        }

        private static bool TryParseSingleStatement(ParserState parser, out string error)
        {
            error = null;
            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                return true;
            }

            if (IsAssignableToken(parser.Current))
            {
                var checkpoint = Save(parser);
                parser.Consume();
                var isAssignment = parser.Current.Kind == SyntaxTokenKind.Operator && parser.Current.Text == "=";
                Restore(parser, checkpoint);
                if (isAssignment)
                {
                    return TryParseAssignmentOrExpression(parser, out error);
                }
            }

            if (parser.Current.Kind == SyntaxTokenKind.Keyword)
            {
                switch (parser.Current.Text)
                {
                    case "REM":
                        parser.Consume();
                        while (parser.Current.Kind != SyntaxTokenKind.End)
                        {
                            parser.Consume();
                        }
                        return true;
                    case "IF":
                        return TryParseIf(parser, out error);
                    case "FOR":
                        return TryParseFor(parser, out error);
                    case "NEXT":
                        return TryParseNext(parser, out error);
                    case "DIM":
                        return TryParseDim(parser, out error);
                    case "POKE":
                        return TryParsePoke(parser, out error);
                    case "GOTO":
                    case "GOSUB":
                    case "GO":
                        return TryParseLineReferenceStatement(parser, out error);
                    case "ON":
                        return TryParseOn(parser, out error);
                    case "DATA":
                        return TryParseData(parser, out error);
                    case "LET":
                        return TryParseGenericKeywordStatement(parser, out error);
                    default:
                        return TryParseGenericKeywordStatement(parser, out error);
                }
            }

            error = "Unknown or invalid statement near '" + parser.Current.Text + "'.";
            return false;
        }

        private static bool TryParseIf(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();

            if (!TryParseExpressionUntil(parser, new string[] { "THEN", "GOSUB" }, out error))
            {
                return false;
            }

            if (parser.Match(SyntaxTokenKind.Keyword, "THEN"))
            {
                if (parser.Current.Kind == SyntaxTokenKind.End)
                {
                    error = "IF statement is missing the THEN clause.";
                    return false;
                }

                if (parser.Current.Kind == SyntaxTokenKind.Number)
                {
                    parser.Consume();
                    return true;
                }

                return TryParseSingleStatement(parser, out error);
            }

            if (parser.Match(SyntaxTokenKind.Keyword, "GOSUB"))
            {
                return TryParseLineReferenceList(parser, out error);
            }

            error = "IF statement is missing THEN or GOSUB.";
            return false;
        }

        private static bool TryParseFor(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();

            if (!TryParseVariableReference(parser, out error))
            {
                error = "FOR statement requires a loop variable.";
                return false;
            }

            if (!parser.Match(SyntaxTokenKind.Operator, "="))
            {
                error = "FOR statement must contain '='.";
                return false;
            }

            if (!TryParseExpressionUntil(parser, new string[] { "TO" }, out error))
            {
                return false;
            }

            if (!parser.Match(SyntaxTokenKind.Keyword, "TO"))
            {
                error = "FOR statement must contain TO.";
                return false;
            }

            if (!TryParseExpressionUntil(parser, new string[] { "STEP" }, out error))
            {
                return false;
            }

            if (parser.Match(SyntaxTokenKind.Keyword, "STEP"))
            {
                return TryParseExpression(parser, out error);
            }

            return true;
        }

        private static bool TryParseNext(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();
            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                return true;
            }

            do
            {
                if (!TryParseVariableReference(parser, out error))
                {
                    error = "NEXT must be followed by variable names.";
                    return false;
                }
            }
            while (parser.Match(SyntaxTokenKind.Comma));

            return true;
        }

        private static bool TryParseDim(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();

            do
            {
                if (!TryParseArrayDeclarator(parser, out error))
                {
                    return false;
                }
            }
            while (parser.Match(SyntaxTokenKind.Comma));

            return true;
        }

        private static bool TryParsePoke(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();

            if (!TryParseExpression(parser, out error))
            {
                return false;
            }

            if (!parser.Match(SyntaxTokenKind.Comma))
            {
                error = "POKE statement must contain a comma.";
                return false;
            }

            return TryParseExpression(parser, out error);
        }

        private static bool TryParseLineReferenceStatement(ParserState parser, out string error)
        {
            parser.Consume();
            return TryParseLineReferenceList(parser, out error);
        }

        private static bool TryParseLineReferenceList(ParserState parser, out string error)
        {
            error = null;
            if (parser.Current.Kind != SyntaxTokenKind.Number)
            {
                error = "Statement should be followed by a line number.";
                return false;
            }

            do
            {
                if (parser.Current.Kind != SyntaxTokenKind.Number)
                {
                    error = "Expected a line number.";
                    return false;
                }

                parser.Consume();
            }
            while (parser.Match(SyntaxTokenKind.Comma));

            return true;
        }

        private static bool TryParseOn(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();

            if (parser.Match(SyntaxTokenKind.Keyword, "ERROR"))
            {
                if (!parser.Match(SyntaxTokenKind.Keyword, "GOTO"))
                {
                    error = "ON ERROR must be followed by GOTO.";
                    return false;
                }

                return TryParseLineReferenceList(parser, out error);
            }

            if (!TryParseExpressionUntil(parser, new string[] { "GOTO", "GOSUB" }, out error))
            {
                return false;
            }

            if (parser.Current.Kind != SyntaxTokenKind.Keyword || (parser.Current.Text != "GOTO" && parser.Current.Text != "GOSUB"))
            {
                error = "ON statement must contain GOTO or GOSUB.";
                return false;
            }

            parser.Consume();
            return TryParseLineReferenceList(parser, out error);
        }

        private static bool TryParseData(ParserState parser, out string error)
        {
            error = null;
            parser.Consume();
            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                return true;
            }

            while (parser.Current.Kind != SyntaxTokenKind.End)
            {
                if (parser.Current.Kind == SyntaxTokenKind.String)
                {
                    parser.Consume();
                }
                else if (!TryParseExpression(parser, out error))
                {
                    return false;
                }

                if (!parser.Match(SyntaxTokenKind.Comma))
                {
                    break;
                }
            }

            return true;
        }

        private static bool TryParseGenericKeywordStatement(ParserState parser, out string error)
        {
            error = null;
            var keyword = parser.Current.Text;
            parser.Consume();

            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                return true;
            }

            if (keyword == "LET")
            {
                return TryParseAssignmentOrExpression(parser, out error);
            }

            if (keyword == "OPTION" && parser.Current.Kind == SyntaxTokenKind.Keyword)
            {
                parser.Consume();
                if (parser.Current.Kind == SyntaxTokenKind.End)
                {
                    return true;
                }
            }

            if ((keyword == "PRINT" || keyword == "CURSOR") && parser.Current.Kind == SyntaxTokenKind.LeftBracket)
            {
                if (!TryParseBracketClause(parser, out error))
                {
                    return false;
                }
            }

            return TryParseDelimitedExpressionList(parser, out error);
        }

        private static bool TryParseAssignmentOrExpression(ParserState parser, out string error)
        {
            error = null;
            if (IsAssignableToken(parser.Current))
            {
                var checkpoint = Save(parser);
                if (TryParseVariableReference(parser, out error))
                {
                    if (parser.Match(SyntaxTokenKind.Operator, "="))
                    {
                        return TryParseExpression(parser, out error);
                    }
                }

                Restore(parser, checkpoint);
            }

            return TryParseDelimitedExpressionList(parser, out error);
        }

        private static bool TryParseDelimitedExpressionList(ParserState parser, out string error)
        {
            error = null;
            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                return true;
            }

            while (parser.Current.Kind != SyntaxTokenKind.End)
            {
                if (parser.Current.Kind == SyntaxTokenKind.Comma || parser.Current.Kind == SyntaxTokenKind.Semicolon)
                {
                    parser.Consume();
                    continue;
                }

                if (!TryParseExpression(parser, out error))
                {
                    return false;
                }

                if (!parser.Match(SyntaxTokenKind.Comma) && !parser.Match(SyntaxTokenKind.Semicolon))
                {
                    break;
                }
            }

            return true;
        }

        private static bool TryParseExpressionUntil(ParserState parser, string[] stopKeywords, out string error)
        {
            error = null;
            if (parser.Current.Kind == SyntaxTokenKind.End)
            {
                error = "Incomplete expression.";
                return false;
            }

            while (parser.Current.Kind != SyntaxTokenKind.End)
            {
                if (parser.Current.Kind == SyntaxTokenKind.Keyword && Array.IndexOf(stopKeywords, parser.Current.Text) >= 0)
                {
                    return parser.Previous.Kind != SyntaxTokenKind.End;
                }

                if (!TryParseExpression(parser, out error))
                {
                    return false;
                }

                if (parser.Current.Kind == SyntaxTokenKind.Keyword && Array.IndexOf(stopKeywords, parser.Current.Text) >= 0)
                {
                    return true;
                }

                if (parser.Current.Kind == SyntaxTokenKind.End)
                {
                    return true;
                }
            }

            return true;
        }

        private static bool TryParseExpression(ParserState parser, out string error)
        {
            return TryParseBinaryExpression(parser, 0, out error);
        }

        private static bool TryParseBinaryExpression(ParserState parser, int minimumPrecedence, out string error)
        {
            error = null;
            if (!TryParseUnary(parser, out error))
            {
                return false;
            }

            while (true)
            {
                var precedence = GetBinaryPrecedence(parser.Current);
                if (precedence < minimumPrecedence)
                {
                    break;
                }

                parser.Consume();
                if (!TryParseBinaryExpression(parser, precedence + 1, out error))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseUnary(ParserState parser, out string error)
        {
            error = null;
            if (parser.Match(SyntaxTokenKind.Operator, "+") || parser.Match(SyntaxTokenKind.Operator, "-") || parser.Match(SyntaxTokenKind.Keyword, "NOT"))
            {
                return TryParseUnary(parser, out error);
            }

            return TryParsePrimary(parser, out error);
        }

        private static bool TryParsePrimary(ParserState parser, out string error)
        {
            error = null;
            if (parser.Match(SyntaxTokenKind.Number) || parser.Match(SyntaxTokenKind.String))
            {
                return true;
            }

            if (parser.Match(SyntaxTokenKind.LeftParen))
            {
                if (!TryParseExpression(parser, out error))
                {
                    return false;
                }

                if (!parser.Match(SyntaxTokenKind.RightParen))
                {
                    error = "Missing closing ')'.";
                    return false;
                }

                return true;
            }

            if (parser.Current.Kind == SyntaxTokenKind.Identifier || parser.Current.Kind == SyntaxTokenKind.Keyword)
            {
                var name = parser.Current.Text;
                parser.Consume();

                if (parser.Match(SyntaxTokenKind.LeftParen))
                {
                    if (!parser.Match(SyntaxTokenKind.RightParen))
                    {
                        do
                        {
                            if (!TryParseExpression(parser, out error))
                            {
                                return false;
                            }
                        }
                        while (parser.Match(SyntaxTokenKind.Comma));

                        if (!parser.Match(SyntaxTokenKind.RightParen))
                        {
                            error = "Missing closing ')' after " + name + ".";
                            return false;
                        }
                    }
                }
                else if (parser.Match(SyntaxTokenKind.LeftBracket))
                {
                    if (!TryParseExpression(parser, out error))
                    {
                        return false;
                    }

                    if (!parser.Match(SyntaxTokenKind.RightBracket))
                    {
                        error = "Missing closing ']'.";
                        return false;
                    }
                }

                return true;
            }

            error = "Expected expression.";
            return false;
        }

        private static bool TryParseVariableReference(ParserState parser, out string error)
        {
            error = null;
            if (!IsAssignableToken(parser.Current))
            {
                return false;
            }

            parser.Consume();
            if (parser.Match(SyntaxTokenKind.LeftParen))
            {
                do
                {
                    if (!TryParseExpression(parser, out error))
                    {
                        return false;
                    }
                }
                while (parser.Match(SyntaxTokenKind.Comma));

                if (!parser.Match(SyntaxTokenKind.RightParen))
                {
                    error = "Missing closing ')' after array reference.";
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseArrayDeclarator(ParserState parser, out string error)
        {
            error = null;
            if (parser.Current.Kind != SyntaxTokenKind.Identifier)
            {
                error = "DIM statement requires an array name.";
                return false;
            }

            parser.Consume();
            if (!parser.Match(SyntaxTokenKind.LeftParen))
            {
                error = "DIM statement is missing array bounds.";
                return false;
            }

            do
            {
                if (!TryParseExpression(parser, out error))
                {
                    return false;
                }
            }
            while (parser.Match(SyntaxTokenKind.Comma));

            if (!parser.Match(SyntaxTokenKind.RightParen))
            {
                error = "DIM statement is missing closing ')'.";
                return false;
            }

            return true;
        }

        private static bool TryParseBracketClause(ParserState parser, out string error)
        {
            error = null;
            if (!parser.Match(SyntaxTokenKind.LeftBracket))
            {
                return true;
            }

            if (!parser.Match(SyntaxTokenKind.RightBracket))
            {
                while (true)
                {
                    if (parser.Current.Kind == SyntaxTokenKind.RightBracket)
                    {
                        parser.Consume();
                        break;
                    }

                    if (parser.Current.Kind != SyntaxTokenKind.Comma)
                    {
                        if (!TryParseExpression(parser, out error))
                        {
                            return false;
                        }
                    }

                    if (parser.Match(SyntaxTokenKind.RightBracket))
                    {
                        break;
                    }

                    if (!parser.Match(SyntaxTokenKind.Comma))
                    {
                        error = "Missing closing ']' in bracket clause.";
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsOperatorStart(string text, int index)
        {
            var ch = text[index];
            return ch == '<' || ch == '>' || ch == '=' || ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^';
        }

        private static string ReadOperator(string text, ref int index)
        {
            if (index + 1 < text.Length)
            {
                var pair = text.Substring(index, 2);
                if (pair == "<>" || pair == "<=" || pair == ">=" || pair == "=<" || pair == "=>" || pair == "><")
                {
                    index += 2;
                    return pair;
                }
            }

            return text[index++].ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryReadNumberToken(string text, ref int index, out string numberText, out string error)
        {
            numberText = null;
            error = null;
            var start = index;

            if (text[index] == '$')
            {
                index++;
                var digitsStart = index;
                while (index < text.Length && IsHexDigit(text[index]))
                {
                    index++;
                }

                if (index == digitsStart)
                {
                    error = "Invalid hex literal after '$'.";
                    return false;
                }

                numberText = text.Substring(start, index - start);
                return true;
            }

            if (text[index] == '&' && index + 1 < text.Length)
            {
                var marker = char.ToUpperInvariant(text[index + 1]);
                if (marker == 'H' || marker == 'O' || marker == 'B')
                {
                    index += 2;
                    var digitsStart = index;
                    while (index < text.Length && IsBaseDigit(text[index], marker))
                    {
                        index++;
                    }

                    if (index == digitsStart)
                    {
                        error = "Invalid base literal after '&" + marker + "'.";
                        return false;
                    }

                    numberText = text.Substring(start, index - start);
                    return true;
                }
            }

            var seenDigit = false;
            var seenDecimal = false;
            while (index < text.Length)
            {
                var ch = text[index];
                if (char.IsDigit(ch))
                {
                    seenDigit = true;
                    index++;
                    continue;
                }

                if (ch == '.' && !seenDecimal)
                {
                    seenDecimal = true;
                    index++;
                    continue;
                }

                break;
            }

            if (!seenDigit)
            {
                error = "Invalid numeric literal.";
                return false;
            }

            numberText = text.Substring(start, index - start);
            return true;
        }

        private static bool IsIdentifierStart(char ch)
        {
            return char.IsLetter(ch);
        }

        private static bool IsIdentifierPart(char ch)
        {
            return char.IsLetterOrDigit(ch) || ch == '$' || ch == '#';
        }

        private static bool IsAssignableToken(SyntaxToken token)
        {
            if (token.Kind == SyntaxTokenKind.Identifier)
            {
                return true;
            }

            return token.Kind == SyntaxTokenKind.Keyword && token.Text.EndsWith("$", StringComparison.Ordinal);
        }

        private static bool TryReadKeywordToken(string text, int index, BasicDialect dialect, out string keyword)
        {
            keyword = null;
            var keywordTokens = GetKeywordTokens(dialect);
            for (var i = 0; i < keywordTokens.Length; i++)
            {
                var candidate = keywordTokens[i];
                if (index + candidate.Length > text.Length)
                {
                    continue;
                }

                if (string.Compare(text, index, candidate, 0, candidate.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    keyword = candidate;
                    return true;
                }
            }

            return false;
        }

        private static int GetBinaryPrecedence(SyntaxToken token)
        {
            if (token.Kind == SyntaxTokenKind.Keyword)
            {
                if (token.Text == "OR")
                {
                    return 1;
                }

                if (token.Text == "AND")
                {
                    return 2;
                }
            }

            if (token.Kind != SyntaxTokenKind.Operator)
            {
                return -1;
            }

            switch (token.Text)
            {
                case "=":
                case "<":
                case ">":
                case "<=":
                case ">=":
                case "=<":
                case "=>":
                case "<>":
                case "><":
                    return 3;
                case "+":
                case "-":
                    return 4;
                case "*":
                case "/":
                    return 5;
                case "^":
                    return 6;
            }

            return -1;
        }

        private static int Save(ParserState parser)
        {
            return parser.Position;
        }

        private static void Restore(ParserState parser, int index)
        {
            parser.Position = index;
        }

        private static bool IsHexDigit(char ch)
        {
            ch = char.ToUpperInvariant(ch);
            return (ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F');
        }

        private static bool IsBaseDigit(char ch, char marker)
        {
            if (marker == 'H')
            {
                return IsHexDigit(ch);
            }

            if (marker == 'O')
            {
                return ch >= '0' && ch <= '7';
            }

            return ch == '0' || ch == '1';
        }

        private static HashSet<string> GetStatementHeads(BasicDialect dialect)
        {
            return dialect == BasicDialect.SBasic ? SBasicStatementHeads : HuBasicStatementHeads;
        }

        private static HashSet<string> GetFunctionNames(BasicDialect dialect)
        {
            return dialect == BasicDialect.SBasic ? SBasicFunctionNames : HuBasicFunctionNames;
        }

        private static string[] GetKeywordTokens(BasicDialect dialect)
        {
            return dialect == BasicDialect.SBasic ? SBasicKeywordTokens : HuBasicKeywordTokens;
        }

        private static HashSet<string> CreateStatementHeads(string[] baseTokens, string[] extTokens)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < baseTokens.Length; i++)
            {
                if (!string.IsNullOrEmpty(baseTokens[i]) && baseTokens[i][0] != '{')
                {
                    set.Add(baseTokens[i]);
                }
            }

            for (var i = 0; i < extTokens.Length; i++)
            {
                if (!string.IsNullOrEmpty(extTokens[i]))
                {
                    set.Add(extTokens[i]);
                }
            }

            return set;
        }

        private static HashSet<string> CreateFunctionNames(string[] functionTokens)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < functionTokens.Length; i++)
            {
                if (!string.IsNullOrEmpty(functionTokens[i]))
                {
                    set.Add(functionTokens[i]);
                }
            }

            return set;
        }

        private static string[] CreateKeywordTokens(HashSet<string> statementHeads, HashSet<string> functionNames)
        {
            var list = new List<string>();
            foreach (var token in statementHeads)
            {
                list.Add(token);
            }

            foreach (var token in functionNames)
            {
                if (!list.Contains(token))
                {
                    list.Add(token);
                }
            }

            if (!list.Contains("AND"))
            {
                list.Add("AND");
            }

            if (!list.Contains("OR"))
            {
                list.Add("OR");
            }

            list.Sort(delegate(string left, string right)
            {
                var compare = right.Length.CompareTo(left.Length);
                return compare != 0 ? compare : string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
            });

            return list.ToArray();
        }
    }

    internal static class BasicEncoder
    {
        private static List<TokenInfo> _sBasicTokens;
        private static List<TokenInfo> _huBasicTokens;

        public static List<BasicLine> ParseEditorText(string text, BasicDialect dialect)
        {
            var result = new List<BasicLine>();
            var sourceLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

            for (var i = 0; i < sourceLines.Length; i++)
            {
                var source = SharpMzTextEncoding.ToLogicalText(sourceLines[i]);
                if (string.IsNullOrWhiteSpace(source))
                {
                    continue;
                }

                var trimmed = source.TrimStart();
                var cursor = 0;
                while (cursor < trimmed.Length && char.IsDigit(trimmed[cursor]))
                {
                    cursor++;
                }

                if (cursor == 0)
                {
                    throw new InvalidDataException(string.Format(CultureInfo.InvariantCulture, "Line {0} is missing a BASIC line number.", i + 1));
                }

                var lineNumber = int.Parse(trimmed.Substring(0, cursor), CultureInfo.InvariantCulture);
                var body = cursor < trimmed.Length ? trimmed.Substring(cursor).TrimStart() : string.Empty;

                result.Add(new BasicLine
                {
                    Number = lineNumber,
                    DisplayBytes = EncodeBodyText(body, dialect).ToArray()
                });
            }

            return result;
        }

        public static byte[] Encode(List<BasicLine> lines, BasicDialect dialect)
        {
            var bytes = new List<byte>();

            for (var i = 0; i < lines.Count; i++)
            {
                var body = lines[i].DisplayBytes ?? new byte[0];
                var lineLength = body.Length + 5;

                bytes.Add((byte)(lineLength & 0xFF));
                bytes.Add((byte)(lineLength >> 8));
                bytes.Add((byte)(lines[i].Number & 0xFF));
                bytes.Add((byte)(lines[i].Number >> 8));
                bytes.AddRange(body);
                bytes.Add(0);
            }

            bytes.Add(0);
            bytes.Add(0);
            return bytes.ToArray();
        }

        private static List<byte> EncodeBodyText(string text, BasicDialect dialect)
        {
            var output = new List<byte>();
            var tokenMap = GetTokens(dialect);
            var inQuotes = false;
            var expectLineReference = false;

            for (var i = 0; i < text.Length;)
            {
                var ch = text[i];

                if (ch == '"')
                {
                    output.Add((byte)'"');
                    inQuotes = !inQuotes;
                    i++;
                    continue;
                }

                if (inQuotes)
                {
                    output.Add(SharpMzTextEncoding.ToSharpByte(ch));
                    i++;
                    continue;
                }

                TokenInfo token;
                if (TryMatchToken(text, i, tokenMap, out token))
                {
                    output.AddRange(token.Bytes);
                    i += token.Text.Length;

                    if (string.Equals(token.Text, "REM", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(token.Text, "DATA", StringComparison.OrdinalIgnoreCase))
                    {
                        while (i < text.Length)
                        {
                            output.Add(SharpMzTextEncoding.ToSharpByte(text[i]));
                            i++;
                        }
                        break;
                    }

                    expectLineReference =
                        string.Equals(token.Text, "GOTO", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(token.Text, "GOSUB", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(token.Text, "THEN", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(token.Text, "RESTORE", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (TryEncodeNumber(text, ref i, dialect, expectLineReference, output))
                {
                    expectLineReference = false;
                    continue;
                }

                output.Add(SharpMzTextEncoding.ToSharpByte(ch));
                if (!char.IsWhiteSpace(ch))
                {
                    expectLineReference = false;
                }
                i++;
            }

            return output;
        }

        private static List<TokenInfo> GetTokens(BasicDialect dialect)
        {
            if (dialect == BasicDialect.SBasic)
            {
                if (_sBasicTokens == null)
                {
                    _sBasicTokens = BuildTokenMap(dialect);
                }

                return _sBasicTokens;
            }

            if (_huBasicTokens == null)
            {
                _huBasicTokens = BuildTokenMap(dialect);
            }

            return _huBasicTokens;
        }

        private static List<TokenInfo> BuildTokenMap(BasicDialect dialect)
        {
            var list = new List<TokenInfo>();
            var baseTokens = dialect == BasicDialect.SBasic ? BasicDecoder.TokenS0 : BasicDecoder.TokenH0;
            var extFe = dialect == BasicDialect.SBasic ? BasicDecoder.TokenSFE : BasicDecoder.TokenHFE;
            var extFf = dialect == BasicDialect.SBasic ? BasicDecoder.TokenSFF : BasicDecoder.TokenHFF;

            AddTokens(list, baseTokens, 0x80, false, 0);
            AddTokens(list, extFe, 0x80, true, 0xFE);
            AddTokens(list, extFf, 0x80, true, 0xFF);

            list.Sort(delegate (TokenInfo a, TokenInfo b)
            {
                return b.Text.Length.CompareTo(a.Text.Length);
            });

            return list;
        }

        private static void AddTokens(List<TokenInfo> list, string[] tokens, int start, bool extended, int prefix)
        {
            for (var i = 0; i < tokens.Length; i++)
            {
                var text = tokens[i];
                if (string.IsNullOrEmpty(text) || text[0] == '{')
                {
                    continue;
                }

                list.Add(new TokenInfo
                {
                    Text = text,
                    Bytes = extended ? new byte[] { (byte)prefix, (byte)(start + i) } : new byte[] { (byte)(start + i) }
                });
            }
        }

        private static bool TryMatchToken(string text, int position, List<TokenInfo> tokenMap, out TokenInfo token)
        {
            for (var i = 0; i < tokenMap.Count; i++)
            {
                var candidate = tokenMap[i];
                if (position + candidate.Text.Length > text.Length)
                {
                    continue;
                }

                if (string.Compare(text, position, candidate.Text, 0, candidate.Text.Length, true, CultureInfo.InvariantCulture) == 0)
                {
                    token = candidate;
                    return true;
                }
            }

            token = null;
            return false;
        }

        private static bool TryEncodeNumber(string text, ref int position, BasicDialect dialect, bool expectLineReference, List<byte> output)
        {
            var start = position;
            if (start >= text.Length)
            {
                return false;
            }

            if (text[position] == '&' && position + 2 < text.Length)
            {
                var marker = char.ToUpperInvariant(text[position + 1]);
                if (marker == 'H' || marker == 'O' || marker == 'B')
                {
                    position += 2;
                    var digitStart = position;
                    while (position < text.Length && IsBaseDigit(text[position], marker))
                    {
                        position++;
                    }

                    if (digitStart == position)
                    {
                        position = start;
                        return false;
                    }

                    var value = ParseBaseValue(text.Substring(digitStart, position - digitStart), marker);
                    output.Add(marker == 'H' ? (dialect == BasicDialect.SBasic ? (byte)0x11 : (byte)0x0F) : marker == 'O' ? (byte)0x0D : (byte)0x0E);
                    output.Add((byte)(value & 0xFF));
                    output.Add((byte)(value >> 8));
                    return true;
                }
            }

            if (text[position] == '$' && position + 1 < text.Length)
            {
                position++;
                var digitStart = position;
                while (position < text.Length)
                {
                    var hexChar = char.ToUpperInvariant(text[position]);
                    if (!((hexChar >= '0' && hexChar <= '9') || (hexChar >= 'A' && hexChar <= 'F')))
                    {
                        break;
                    }

                    position++;
                }

                if (digitStart == position)
                {
                    position = start;
                    return false;
                }

                var value = ParseBaseValue(text.Substring(digitStart, position - digitStart), 'H');
                output.Add(dialect == BasicDialect.SBasic ? (byte)0x11 : (byte)0x0F);
                output.Add((byte)(value & 0xFF));
                output.Add((byte)(value >> 8));
                return true;
            }

            if (!char.IsDigit(text[position]))
            {
                return false;
            }

            while (position < text.Length && char.IsDigit(text[position]))
            {
                position++;
            }

            var hasDecimal = false;
            if (position < text.Length && text[position] == '.')
            {
                hasDecimal = true;
                position++;
                while (position < text.Length && char.IsDigit(text[position]))
                {
                    position++;
                }
            }

            if (position < text.Length && (text[position] == 'E' || text[position] == 'e'))
            {
                hasDecimal = true;
                position++;
                if (position < text.Length && (text[position] == '+' || text[position] == '-'))
                {
                    position++;
                }
                while (position < text.Length && char.IsDigit(text[position]))
                {
                    position++;
                }
            }

            var tokenText = text.Substring(start, position - start);
            if (expectLineReference)
            {
                ushort lineRef;
                if (!ushort.TryParse(tokenText, NumberStyles.Integer, CultureInfo.InvariantCulture, out lineRef))
                {
                    throw new InvalidDataException("Line reference is out of range.");
                }

                output.Add(0x0B);
                output.Add((byte)(lineRef & 0xFF));
                output.Add((byte)(lineRef >> 8));
                return true;
            }

            if (!hasDecimal)
            {
                int integerValue;
                if (int.TryParse(tokenText, NumberStyles.Integer, CultureInfo.InvariantCulture, out integerValue))
                {
                    if (dialect == BasicDialect.HuBasic && integerValue >= 0 && integerValue <= 9)
                    {
                        output.Add((byte)(integerValue + 1));
                        return true;
                    }

                    if (dialect == BasicDialect.HuBasic && integerValue >= short.MinValue && integerValue <= short.MaxValue)
                    {
                        output.Add(0x12);
                        var word = unchecked((ushort)(short)integerValue);
                        output.Add((byte)(word & 0xFF));
                        output.Add((byte)(word >> 8));
                        return true;
                    }
                }
            }

            double floatValue;
            if (!double.TryParse(tokenText, NumberStyles.Float, CultureInfo.InvariantCulture, out floatValue))
            {
                position = start;
                return false;
            }

            output.Add(0x15);
            output.AddRange(EncodeFloat(floatValue));
            return true;
        }

        private static bool IsBaseDigit(char ch, char marker)
        {
            if (marker == 'H')
            {
                return (ch >= '0' && ch <= '9') || (char.ToUpperInvariant(ch) >= 'A' && char.ToUpperInvariant(ch) <= 'F');
            }

            if (marker == 'O')
            {
                return ch >= '0' && ch <= '7';
            }

            return ch == '0' || ch == '1';
        }

        private static ushort ParseBaseValue(string text, char marker)
        {
            var value = 0;
            for (var i = 0; i < text.Length; i++)
            {
                int digit;
                var ch = char.ToUpperInvariant(text[i]);
                if (ch >= '0' && ch <= '9')
                {
                    digit = ch - '0';
                }
                else
                {
                    digit = 10 + (ch - 'A');
                }

                value = value * (marker == 'H' ? 16 : marker == 'O' ? 8 : 2) + digit;
            }

            return (ushort)value;
        }

        private static byte[] EncodeFloat(double value)
        {
            if (value == 0.0)
            {
                return new byte[] { 0, 0, 0, 0, 0 };
            }

            var signBit = value < 0 ? 0x80000000U : 0U;
            var absoluteValue = Math.Abs(value);
            var exponent = (int)Math.Floor(Math.Log(absoluteValue, 2.0)) + 1;
            var scaled = absoluteValue / Math.Pow(2.0, exponent - 32);
            var mantissa = (uint)Math.Round(scaled, MidpointRounding.AwayFromZero);

            if (mantissa == 0xFFFFFFFF)
            {
                mantissa >>= 1;
                exponent++;
            }

            var raw = (mantissa & 0x7FFFFFFF) | signBit;
            return new byte[]
            {
                (byte)(exponent + 128),
                (byte)(raw >> 24),
                (byte)(raw >> 16),
                (byte)(raw >> 8),
                (byte)raw
            };
        }
    }

    internal sealed class BasicProgramView : FrameworkElement
    {
        private const int GlyphIndexOffset = 1;
        private const double MarginSize = 24.0;
        private const double LineHeight = 24.0;
        private const double FontRenderingSize = 20.0;

        private List<BasicLine> _lines = new List<BasicLine>();
        private double _cellWidth = 14.0;
        private int _maxColumns = 1;

        public GlyphTypeface GlyphTypeface { get; set; }

        public Brush BackgroundBrush { get; set; }

        public Brush ForegroundBrush { get; set; }

        public Brush AccentBrush { get; set; }

        public void SetProgram(List<BasicLine> lines)
        {
            _lines = lines ?? new List<BasicLine>();
            RecalculateMetrics();
            InvalidateMeasure();
            InvalidateVisual();
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            return GetDesiredSize();
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            return finalSize;
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);

            var desiredSize = GetDesiredSize();
            drawingContext.DrawRectangle(BackgroundBrush ?? Brushes.Black, null, new Rect(0, 0, desiredSize.Width, desiredSize.Height));

            DrawHeader(drawingContext, desiredSize.Width);

            var foreground = ForegroundBrush ?? Brushes.White;
            for (var lineIndex = 0; lineIndex < _lines.Count; lineIndex++)
            {
                var line = _lines[lineIndex];
                var y = MarginSize + 28.0 + (lineIndex * LineHeight);
                DrawLineNumber(drawingContext, line.Number, y);
                DrawBytes(drawingContext, line.DisplayBytes, MarginSize + (_cellWidth * 7.5), y, foreground);
            }

            if (_lines.Count == 0)
            {
                DrawPlaceholder(drawingContext, desiredSize.Width, desiredSize.Height);
            }
        }

        private Size GetDesiredSize()
        {
            var width = MarginSize * 2.0 + (_cellWidth * (_maxColumns + 10));
            var height = MarginSize * 2.0 + 28.0 + (_lines.Count * LineHeight);
            if (height < 220)
            {
                height = 220;
            }

            return new Size(width, height);
        }

        private void RecalculateMetrics()
        {
            _maxColumns = 1;

            if (GlyphTypeface != null && GlyphTypeface.AdvanceWidths.ContainsKey(0x21))
            {
                _cellWidth = GlyphTypeface.AdvanceWidths[0x21] * FontRenderingSize;
            }
            else
            {
                _cellWidth = 14.0;
            }

            foreach (var line in _lines)
            {
                if (line.DisplayBytes != null && line.DisplayBytes.Length > _maxColumns)
                {
                    _maxColumns = line.DisplayBytes.Length;
                }
            }
        }

        private void DrawHeader(DrawingContext drawingContext, double width)
        {
            var pen = new Pen(AccentBrush ?? Brushes.Orange, 2);
            drawingContext.DrawLine(pen, new Point(MarginSize, MarginSize + 18), new Point(width - MarginSize, MarginSize + 18));
        }

        private void DrawLineNumber(DrawingContext drawingContext, int number, double y)
        {
            var text = number.ToString(CultureInfo.InvariantCulture).PadLeft(5) + " ";
            DrawAsciiText(drawingContext, text, MarginSize, y, AccentBrush ?? Brushes.Orange);
        }

        private void DrawBytes(DrawingContext drawingContext, byte[] bytes, double x, double y, Brush brush)
        {
            if (GlyphTypeface == null || bytes == null)
            {
                return;
            }

            for (var i = 0; i < bytes.Length; i++)
            {
                DrawGlyph(drawingContext, bytes[i], x + (_cellWidth * i), y, brush);
            }
        }

        private void DrawAsciiText(DrawingContext drawingContext, string text, double x, double y, Brush brush)
        {
            var bytes = Encoding.ASCII.GetBytes(text);
            DrawBytes(drawingContext, bytes, x, y, brush);
        }

        private void DrawGlyph(DrawingContext drawingContext, byte code, double x, double y, Brush brush)
        {
            var glyphIndex = (ushort)(code + GlyphIndexOffset);
            double advance = _cellWidth;
            if (GlyphTypeface.AdvanceWidths.ContainsKey(glyphIndex))
            {
                advance = GlyphTypeface.AdvanceWidths[glyphIndex] * FontRenderingSize;
            }

            var origin = new Point(x, y + 18.0);
            var glyphRun = new GlyphRun(
                GlyphTypeface,
                0,
                false,
                FontRenderingSize,
                new ushort[] { glyphIndex },
                origin,
                new double[] { advance },
                null,
                null,
                null,
                null,
                null,
                null);

            drawingContext.DrawGlyphRun(brush, glyphRun);
        }

        private void DrawPlaceholder(DrawingContext drawingContext, double width, double height)
        {
            var brush = new SolidColorBrush(Color.FromRgb(0xC2, 0xC9, 0xCF));
            DrawAsciiText(drawingContext, "OPEN A SHARP MZ BASIC .MZF FILE", MarginSize, height / 2.0, brush);
        }
    }

    internal static class BasicDecoder
    {
        internal static readonly string[] TokenS0 = new string[]
        {
            "GOTO", "GOSUB", null, "RUN", "RETURN", "RESTORE", "RESUME", "LIST",
            null, "DELETE", "RENUM", "AUTO", null, "FOR", "NEXT", "PRINT",
            null, "INPUT", null, "IF", "DATA", "READ", "DIM", "REM",
            "END", "STOP", "CONT", "CLS", null, "ON", "LET", "NEW",
            "POKE", "OFF", "MODE", "SKIP", "PLOT", "LINE", "RLINE", "MOVE",
            "RMOVE", "TRON", "TROFF", "INP#", null, "GET", "PCOLOR", "PHOME",
            "HSET", "GPRINT", "KEY", "AXIS", "LOAD", "SAVE", "MERGE", null,
            "CONSOLE", null, "OUT#", "CIRCLE", "TEST", "PAGE", null, null,
            "ERASE", "ERROR", null, "USR", "BYE", null, null, "DEF",
            null, null, null, null, null, null, "WOPEN", "CLOSE",
            "ROPEN", null, null, null, null, null, null, null,
            null, "KILL", null, null, null, null, null, null,
            "TO", "STEP", "THEN", "USING", "{E4}", null, "TAB", "SPC",
            null, null, null, "OR", "AND", null, "><", "<>",
            "=<", "<=", "=>", ">=", "=", ">", "<", "+",
            "-", null, null, "/", "*", "^"
        };

        internal static readonly string[] TokenSFE = new string[]
        {
            null, "SET", "RESET", "COLOR", null, null, null, null,
            null, null, null, null, null, null, null, null,
            null, null, null, null, null, null, null, null,
            null, null, null, null, null, null, null, null,
            null, null, "MUSIC", "TEMPO", "CURSOR", "VERIFY", "CLR", "LIMIT",
            null, null, null, null, null, null, "BOOT", null
        };

        internal static readonly string[] TokenSFF = new string[]
        {
            "INT", "ABS", "SIN", "COS", "TAN", "LN", "EXP", "SQR",
            "RND", "PEEK", "ATN", "SGN", "LOG", null, "PAI", "RAD",
            null, null, null, null, null, "EOF", null, null,
            null, null, null, null, null, null, "JOY", null,
            "CHR$", "STR$", "HEX$", null, null, null, null, null,
            null, null, null, "ASC", "LEN", "VAL", null, null,
            null, null, null, "ERN", "ERL", "SIZE", null, null,
            null, null, "LEFT$", "RIGHT$", "MID$", null, null, null,
            null, null, null, "STRING$", "TI$", null, null, "FN",
            null, null, null, null, null, null, null, null,
            null
        };

        internal static readonly string[] TokenH0 = new string[]
        {
            "GOTO", "GOSUB", "GO", "RUN", "RETURN", "RESTORE", "RESUME", "LIST",
            "LLIST", "DELETE", "RENUM", "AUTO", "EDIT", "FOR", "NEXT", "PRINT",
            "LPRINT", "INPUT", "LINPUT", "IF", "DATA", "READ", "DIM", "REM",
            "END", "STOP", "CONT", "CLS", "CLEAR", "ON", "LET", "NEW",
            "POKE", "OFF", "WHILE", "WEND", "REPEAT", "UNTIL", null, null,
            "TRACE", "TRON", "TROFF", "SPEED", null, null, "DEFINT", "DEFSNG",
            "DEFDBL", "DEFSTR", "DEF", null, "LOAD", "SAVE", "MERGE", "CHAIN",
            "CONSOLE", null, "OUT", "SEARCH", "WAIT", "PAUSE", "WRITE", "SWAP",
            "ERASE", "ERROR", "ELSE", "CALL", "MON", "LOCATE", "MODE", "KEY",
            "PUSH", "POP", "LABEL", "RANDOMIZE", "OPTION", "LINE", "OPEN", "CLOSE",
            null, "FIELD", "GET", "PUT", "SET", "FILES", "LFILES", "DEVICE",
            "NAME", "KILL", "LSET", "RSET", "INIT", "VDIM", "MAXFILES", null,
            "TO", "STEP", "THEN", "USING", "SUB", "BASE", "TAB", "SPC",
            "EQV", "IMP", "XOR", "OR", "AND", "NOT", "><", "<>",
            "=<", "<=", "=>", ">=", "=", ">", "<", "+",
            "-", "MOD", "\\", "/", "*", "^"
        };

        internal static readonly string[] TokenHFE = new string[]
        {
            null, "PSET", "PRESET", "COLOR", null, null, null, null,
            null, null, null, "PLAY", null, "BEEP", null, null,
            null, null, null, null, "CGEN", "PCOLOR", "SKIP", "RLINE",
            "MOVE", "RMOVE", "PHOME", "HSET", "GPRINT", "AXIS", "CIRCLE", "TEST",
            "PLOT", "PAGE", "MUSIC", "TEMPO", "CURSOR", "VERIFY", "CLR", "LIMIT",
            "KLIST", null, null, "CLICK", "BOOT", "DEVI$", "DEVO$", null
        };

        internal static readonly string[] TokenHFF = new string[]
        {
            "INT", "ABS", "SIN", "COS", "TAN", "LOG", "EXP", "SQR",
            "RND", "PEEK", "ATN", "SGN", "FRAC", "FIX", "PAI", "RAD",
            "INP", "CDBL", "CSNG", "CINT", "DSKF", "EOF", "FPOS", "LOC",
            "LOF", "POS", "FAC", "SUM", "FRE", null, "JOY", null,
            "CHR$", "STR$", "HEX$", "OCT$", "BIN$", "MKI$", "MKS$", "MKD$",
            "SPACE$", null, null, "ASC", "LEN", "VAL", "CVS", "CVD",
            "CVI", null, null, "ERR", "ERL", "CSRLIN", "STRPTR", "DTL",
            null, null, "LEFT$", "RIGHT$", "MID$", "INKEY$", "INSTR", "HEXCHR$",
            "MEM$", "SCRN$", "VARPTR", "STRING$", "TIME$", null, null, "FN",
            "USR", null, null, "ATTR$", null, "CHARACTER$", null
        };

        public static List<BasicLine> Decode(byte[] bytes, BasicDialect dialect)
        {
            if (bytes == null || bytes.Length <= 128)
            {
                throw new InvalidDataException("The file is too small to contain an MZF BASIC payload.");
            }

            var reader = new ByteReader(bytes, 128);
            var lines = new List<BasicLine>();

            while (true)
            {
                var lineLength = reader.GetWord();
                if (lineLength == 0)
                {
                    break;
                }

                if ((reader.Position + lineLength - 3) >= bytes.Length || bytes[reader.Position + lineLength - 3] != 0)
                {
                    throw new InvalidDataException("Invalid BASIC line structure in the MZF payload.");
                }

                var lineNumber = reader.GetWord();
                var output = new List<byte>();

                while (true)
                {
                    var c = reader.GetByte();
                    if (c == 0)
                    {
                        break;
                    }

                    if ((c & 0x80) != 0)
                    {
                        ExpandToken(reader, output, c, lineNumber, dialect);
                        if (c == 0x94 || c == 0x97)
                        {
                            SkipData(reader, output);
                        }
                    }
                    else
                    {
                        ExpandPlainByte(reader, output, c, dialect);
                    }
                }

                lines.Add(new BasicLine
                {
                    Number = lineNumber,
                    DisplayBytes = output.ToArray()
                });
            }

            return lines;
        }

        private static void ExpandToken(ByteReader reader, List<byte> output, byte c, int lineNumber, BasicDialect dialect)
        {
            string token = null;

            if (c == 0xFE)
            {
                var next = reader.GetByte();
                if (dialect == BasicDialect.SBasic)
                {
                    if (next > 0xAE) next = 0xAF;
                    token = TokenSFE[next & 0x7F];
                }
                else
                {
                    if (next > 0xAE) next = 0xAF;
                    token = TokenHFE[next & 0x7F];
                }
            }
            else if (c == 0xFF)
            {
                var next = reader.GetByte();
                if (dialect == BasicDialect.SBasic)
                {
                    if (next > 0xC7) next = 0xC8;
                    token = TokenSFF[next & 0x7F];
                }
                else
                {
                    if (next > 0xCD) next = 0xCE;
                    token = TokenHFF[next & 0x7F];
                }
            }
            else
            {
                token = dialect == BasicDialect.SBasic ? TokenS0[c & 0x7F] : TokenH0[c & 0x7F];
            }

            if (token == null)
            {
                throw new InvalidDataException(string.Format(CultureInfo.InvariantCulture, "Syntax error in line {0} near token {1:X2}.", lineNumber, c));
            }

            AppendTokenText(output, token);
        }

        private static void ExpandPlainByte(ByteReader reader, List<byte> output, byte c, BasicDialect dialect)
        {
            if (dialect == BasicDialect.HuBasic && c >= 1 && c <= 10)
            {
                output.Add((byte)(0x30 - 1 + c));
                return;
            }

            switch (c)
            {
                case 0x0B:
                    AppendAscii(output, reader.GetWord().ToString(CultureInfo.InvariantCulture));
                    break;
                case 0x0D:
                    AppendAscii(output, "&O" + Convert.ToString(reader.GetWord(), 8));
                    break;
                case 0x0E:
                    AppendAscii(output, "&B" + Convert.ToString(reader.GetWord(), 2));
                    break;
                case 0x0F:
                    AppendAscii(output, "&H" + reader.GetWord().ToString("X", CultureInfo.InvariantCulture));
                    break;
                case 0x11:
                    AppendAscii(output, "$" + reader.GetWord().ToString("X", CultureInfo.InvariantCulture));
                    break;
                case 0x12:
                    AppendAscii(output, unchecked((short)reader.GetWord()).ToString(CultureInfo.InvariantCulture));
                    break;
                case 0x15:
                    AppendAscii(output, FormatNumber(reader.GetFloat()));
                    break;
                case 0x18:
                    AppendAscii(output, FormatNumber(reader.GetDouble()));
                    break;
                case 0x3A:
                    output.Add(c);
                    if (dialect == BasicDialect.HuBasic)
                    {
                        var peek = reader.PeekByte();
                        if (peek == 0x27)
                        {
                            SkipData(reader, output);
                        }
                    }
                    break;
                case 0x22:
                    output.Add(c);
                    SkipString(reader, output);
                    break;
                default:
                    output.Add(c);
                    break;
            }
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("G", CultureInfo.InvariantCulture);
        }

        private static void SkipData(ByteReader reader, List<byte> output)
        {
            var inQuotes = false;
            while (true)
            {
                var c = reader.GetByte();
                if (!inQuotes && c == 0)
                {
                    reader.StepBack();
                    return;
                }

                output.Add(c);
                if (c == 0x22)
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (!inQuotes && c == 0x3A)
                {
                    return;
                }
            }
        }

        private static void SkipString(ByteReader reader, List<byte> output)
        {
            while (true)
            {
                var c = reader.GetByte();
                if (c == 0)
                {
                    reader.StepBack();
                    return;
                }

                output.Add(c);
                if (c == 0x22)
                {
                    return;
                }
            }
        }

        private static void AppendTokenText(List<byte> output, string text)
        {
            if (text.Length == 4 && text[0] == '{' && text[3] == '}')
            {
                byte value;
                if (byte.TryParse(text.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value))
                {
                    output.Add(value);
                    return;
                }
            }

            AppendAscii(output, text);
        }

        private static void AppendAscii(List<byte> output, string text)
        {
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                output.Add(ch <= 0xFF ? (byte)ch : (byte)'?');
            }
        }
    }

    internal sealed class ByteReader
    {
        private readonly byte[] _buffer;
        private int _position;

        public ByteReader(byte[] buffer, int start)
        {
            _buffer = buffer;
            _position = start;
        }

        public int Position
        {
            get { return _position; }
        }

        public byte GetByte()
        {
            EnsureAvailable(1);
            return _buffer[_position++];
        }

        public byte PeekByte()
        {
            EnsureAvailable(1);
            return _buffer[_position];
        }

        public ushort GetWord()
        {
            EnsureAvailable(2);
            var value = (ushort)(_buffer[_position] | (_buffer[_position + 1] << 8));
            _position += 2;
            return value;
        }

        public double GetFloat()
        {
            var e = GetByte();
            EnsureAvailable(4);
            var raw =
                ((uint)_buffer[_position] << 24) |
                ((uint)_buffer[_position + 1] << 16) |
                ((uint)_buffer[_position + 2] << 8) |
                _buffer[_position + 3];
            _position += 4;

            var sign = (raw & 0x80000000U) != 0 ? -1.0 : 1.0;
            if (e == 0)
            {
                return 0.0;
            }

            var exponent = e - 128;
            return sign * (raw | 0x80000000U) * Math.Pow(2.0, exponent - 32);
        }

        public double GetDouble()
        {
            var e = GetByte();
            EnsureAvailable(7);
            var rh =
                ((uint)_buffer[_position] << 24) |
                ((uint)_buffer[_position + 1] << 16) |
                ((uint)_buffer[_position + 2] << 8) |
                _buffer[_position + 3];
            var rl =
                ((uint)_buffer[_position + 4] << 24) |
                ((uint)_buffer[_position + 5] << 16) |
                ((uint)_buffer[_position + 6] << 8);
            _position += 7;

            var sign = (rh & 0x80000000U) != 0 ? -1.0 : 1.0;
            if (e == 0)
            {
                return 0.0;
            }

            var exponent = e - 128;
            var value = (rh | 0x80000000U) * Math.Pow(2.0, exponent - 32);
            value += rl * Math.Pow(2.0, exponent - 56);
            return sign * value;
        }

        public void StepBack()
        {
            if (_position > 0)
            {
                _position--;
            }
        }

        private void EnsureAvailable(int byteCount)
        {
            if (_position + byteCount > _buffer.Length)
            {
                throw new EndOfStreamException("Unexpected end of BASIC data.");
            }
        }
    }
}

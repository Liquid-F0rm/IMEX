using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using ImageMagick;
using Microsoft.Win32;
using SkiaSharp;
using System.Drawing; // Bitmap
using System.Runtime.InteropServices;
using Windows.Globalization;
using Windows.Media.Ocr;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;
using WpfBitmapFrame = System.Windows.Media.Imaging.BitmapFrame;
using WinBitmapDecoder = Windows.Graphics.Imaging.BitmapDecoder;
using DrawingPoint = System.Drawing.Point;
using WpfPoint = System.Windows.Point;

namespace WpfImageToText
{
    // Structures pour l'API Windows
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    public partial class MainWindow : Window
    {
#pragma warning disable CA1416 // Valider la compatibilité de la plateforme
        private readonly string tempFolder;
        private readonly string autosaveFolderPath;
        private readonly string autosaveImagePath;
        private readonly string autosaveTextPath;
        private readonly string autosaveConfigPath;
        public object? UnsharpMask { get; private set; }
        private bool isDarkTheme = false;
        private bool isEnglish = false; // Par défaut, l'interface est en français

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr hwnd,
            int attr,
            ref int attrValue,
            int attrSize);

        private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        private const int DWMWCP_ROUND = 2;

        public new double MinWidth { get; set; } = 400;
        public new double MinHeight { get; set; } = 300;

        public MainWindow()
        {
            InitializeComponent();

            this.ResizeMode = ResizeMode.CanResizeWithGrip;
            UnsharpMask = null;

            tempFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WpfImageToText");
            if (!Directory.Exists(tempFolder))
                Directory.CreateDirectory(tempFolder);

            // Définir les chemins d'autosauvegarde
            autosaveFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WpfImageToText");
            autosaveImagePath = Path.Combine(autosaveFolderPath, "autosave_image.png");
            autosaveTextPath = Path.Combine(autosaveFolderPath, "autosave_text.txt");
            autosaveConfigPath = Path.Combine(autosaveFolderPath, "autosave_config.txt");
            
            // Créer le dossier s'il n'existe pas
            if (!Directory.Exists(autosaveFolderPath))
                Directory.CreateDirectory(autosaveFolderPath);

            CleanupOldTempFiles();

            Loaded += (s, e) =>
            {
                var settings = Properties.Settings.Default;
                if (!string.IsNullOrEmpty(settings.Paramètre))
                {
                    Width = ExtractNumericValue(settings.Paramètre, 900);
                    Height = ExtractNumericValue(settings.Paramètre1, 600);
                    Left = ExtractNumericValue(settings.Paramètre2, 100);
                    Top = ExtractNumericValue(settings.Paramètre3, 100);
                }

                string themeValue = settings.ThemeParameter ?? "light";
                isDarkTheme = themeValue.Equals("dark", StringComparison.OrdinalIgnoreCase);
                ApplyTheme();

                // Charger la préférence de langue
                string languageValue = settings.LanguageParameter ?? "fr";
                isEnglish = languageValue.Equals("en", StringComparison.OrdinalIgnoreCase);

                // Mettre à jour le texte du bouton de langue
                LanguageButton.Content = isEnglish ? "English" : "Français";

                // Appliquer la langue chargée
                ApplyLanguage();

                // Restaurer la session précédente
                RestoreLastSession();

                IntPtr hWnd = new WindowInteropHelper(this).Handle;
                int preference = DWMWCP_ROUND;
                DwmSetWindowAttribute(hWnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref preference, sizeof(int));
            };
        }

        protected override void OnClosed(EventArgs e)
        {
            var settings = Properties.Settings.Default;
            settings.Paramètre = Width.ToString();
            settings.Paramètre1 = Height.ToString();
            settings.Paramètre2 = Left.ToString();
            settings.Paramètre3 = Top.ToString();

            settings.ThemeParameter = isDarkTheme ? "dark" : "light";
            settings.Save();

            base.OnClosed(e);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var hwndSource = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSource != null)
            {
                hwndSource.AddHook(WindowProc);
            }
        }

        private IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_GETMINMAXINFO = 0x0024;
            if (msg == WM_GETMINMAXINFO)
            {
                MINMAXINFO mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
                mmi.ptMinTrackSize.x = (int)MinWidth;
                mmi.ptMinTrackSize.y = (int)MinHeight;
                Marshal.StructureToPtr(mmi, lParam, false);
                handled = false;
            }
            return IntPtr.Zero;
        }

        #region Browse / Drop / Preview

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.tiff";
            if (dlg.ShowDialog() == true)
            {
                LoadImage(dlg.FileName);
            }
        }

        private void ImageDrop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                    LoadImage(files[0]);
            }
        }

        private void Border_PreviewDragOver(object sender, System.Windows.DragEventArgs e)
        {
            e.Handled = true;
        }

        private async void LoadImage(string path)
        {
            try
            {
                // Supprimer les fichiers d'autosauvegarde existants
                DeleteAutosaveSession();
                
                ImagePreview.Source = null;
                OcrTextBox.Clear();

                string extension = System.IO.Path.GetExtension(path).ToLowerInvariant();
                if (extension != ".png" && extension != ".jpg" && extension != ".jpeg" && extension != ".bmp" && extension != ".tiff")
                    throw new InvalidDataException(GetLocalizedMessage("UnsupportedFileFormat"));

                string tempFileName = "img_" + Guid.NewGuid().ToString() + extension;
                string tempPath = System.IO.Path.Combine(tempFolder, tempFileName);
                File.Copy(path, tempPath, true);

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(tempPath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                ImagePreview.Source = bitmap;

                // Aperçu redimensionné (optionnel)
                using (var originalBitmap = new Bitmap(tempPath))
                using (var resizedBitmap = new Bitmap(originalBitmap, new System.Drawing.Size(400, 300))
                )
                {
                    string resizedTempPath = System.IO.Path.Combine(tempFolder, "resized_" + tempFileName);
                    resizedBitmap.Save(resizedTempPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                // Lancer l'OCR
                try
                {
                    await RunWindowsOcr(tempPath);
                }
                catch (Exception ocrEx)
                {
                    OcrTextBox.Text = GetLocalizedMessage("OcrExecutionError") + ocrEx.Message;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetLocalizedMessage("ImageLoadError") + ex.Message);
            }
        }

        // Ajouter cette méthode pour supprimer les fichiers d'autosauvegarde
        private void DeleteAutosaveSession()
        {
            try
            {
                // Supprimer les fichiers d'autosauvegarde s'ils existent
                if (File.Exists(autosaveImagePath))
                    File.Delete(autosaveImagePath);
                    
                if (File.Exists(autosaveTextPath))
                    File.Delete(autosaveTextPath);
                    
                if (File.Exists(autosaveConfigPath))
                    File.Delete(autosaveConfigPath);
            }
            catch (Exception ex)
            {
                // En cas d'erreur, nous continuons silencieusement - pas besoin d'alarmer l'utilisateur
                File.AppendAllText(Path.Combine(tempFolder, "delete_session_error_log.txt"),
                    $"[{DateTime.Now}] Erreur suppression session: {ex.Message}\n{ex.StackTrace}\n\n");
            }
        }

        #endregion

        // ---------------- OCR: Prétraitement amélioré ----------------
        private string PreprocessImageForOcr(string originalImagePath)
        {
            try
            {
                string processedImagePath = System.IO.Path.Combine(tempFolder,
                    "win_" + System.IO.Path.GetFileName(originalImagePath));

                using (var bitmap = SKBitmap.Decode(originalImagePath))
                {
                    if (bitmap == null)
                        return originalImagePath;

                    bool isLowRes = Math.Max(bitmap.Width, bitmap.Height) < 900;

                    float scale = 1.0f;
                    if (isLowRes)
                    {
                        // gentle upscale factor depending on size
                        int maxSide = Math.Max(bitmap.Width, bitmap.Height);
                        scale = maxSide < 600 ? 2.5f : 1.8f; // stronger upscaling for very small images
                    }

                    int targetWidth = (int)Math.Max(1, bitmap.Width * scale);
                    int targetHeight = (int)Math.Max(1, bitmap.Height * scale);

                    var info = new SKImageInfo(targetWidth, targetHeight, bitmap.ColorType, bitmap.AlphaType, bitmap.ColorSpace);
                    using (var surface = SKSurface.Create(info))
                    {
                        var canvas = surface.Canvas;
                        canvas.Clear(SKColors.Transparent);

                        // Version corrigée avec la nouvelle API SkiaSharp
                        using (var paint = new SKPaint { IsAntialias = true })
                        {
                            // Dessiner l'image mise à l'échelle
                            canvas.DrawBitmap(bitmap, new SKRect(0, 0, targetWidth, targetHeight), paint);
                        }

                        // compute average luminance (sampled)
                        double sumLuma = 0;
                        int samples = 0;
                        int stepX = Math.Max(1, bitmap.Width / 20);
                        int stepY = Math.Max(1, bitmap.Height / 20);
                        for (int y = 0; y < bitmap.Height; y += stepY)
                        {
                            for (int x = 0; x < bitmap.Width; x += stepX)
                            {
                                var c = bitmap.GetPixel(x, y);
                                sumLuma += 0.2126 * c.Red + 0.7152 * c.Green + 0.0722 * c.Blue;
                                samples++;
                            }
                        }
                        double avgLuma = samples > 0 ? sumLuma / samples / 255.0 : 0.5;
                        bool isDarkBackground = avgLuma < 0.35;

                        // Windows OCR: keep transformations light
                        using (var gammaPaint = new SKPaint())
                        {
                            gammaPaint.ColorFilter = SKColorFilter.CreateLighting(new SKColor(250, 250, 250), new SKColor(8, 8, 8));
                            // Dessiner avec le filtre gamma
                            canvas.DrawBitmap(bitmap, new SKRect(0, 0, targetWidth, targetHeight), gammaPaint);
                        }

                        using (var img = surface.Snapshot())
                        using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                        using (var fs = File.Open(processedImagePath, FileMode.Create, FileAccess.Write))
                        {
                            data.SaveTo(fs);
                        }

                        return processedImagePath;
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(System.IO.Path.Combine(tempFolder, "preprocessing_error_log.txt"),
                    $"[{DateTime.Now}] Erreur prétraitement: {ex.Message}\n{ex.StackTrace}\n\n");
                return originalImagePath;
            }
        }

        // ---------------- Windows OCR ----------------
        private async Task RunWindowsOcr(string originalImagePath, string variant = "moyenne")
        {
            try
            {
                // Créer un chemin temporaire unique
                string tempFolderPath = Path.Combine(Path.GetTempPath(), "WpfOcrApp");
                Directory.CreateDirectory(tempFolderPath);

                string processedImagePath = Path.Combine(tempFolderPath, $"win_{Guid.NewGuid():N}.png");

                // Prétraitement selon la variante
                switch (variant.ToLower())
                {
                    case "douce":
                        PreprocessImageForOcr_Douce(originalImagePath, processedImagePath);
                        break;
                    case "moyenne":
                        PreprocessImageForOcr_Moyenne(originalImagePath, processedImagePath);
                        break;
                    case "agressive":
                        PreprocessImageForOcr_Agressive(originalImagePath, processedImagePath);
                        break;
                    default:
                        PreprocessImageForOcr_Moyenne(originalImagePath, processedImagePath);
                        break;
                }

                if (!File.Exists(processedImagePath))
                {
                    OcrTextBox.Text = GetLocalizedMessage("PreprocessedFileNotFound");
                    return;
                }

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.UriSource = new Uri(processedImagePath);
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                var softwareBitmap = await ConvertBitmapImageToSoftwareBitmap(bitmapImage);
                if (softwareBitmap == null)
                {
                    OcrTextBox.Text = GetLocalizedMessage("ImageConversionFailed");
                    return;
                }

                OcrEngine? ocrEngine = null;
                string resultText = "";

                if (isEnglish)
                {
                    // Comportement inchangé pour l'interface en anglais
                    ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("en-US"));
                    if (ocrEngine == null)
                    {
                        OcrTextBox.Text = GetLocalizedMessage("WindowsOcrNotAvailable") + "\n\n" + GetLocalizedMessage("OcrLanguagePackInstruction");
                        return;
                    }
                }
                else
                {
                    // Pour l'interface en français:
                    // 1. Essayer d'abord l'OCR anglais
                    ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("en-US"));
                    
                    // 2. Si l'OCR anglais n'est pas disponible, utiliser l'OCR français sans alerter
                    if (ocrEngine == null)
                    {
                        ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("fr-FR"));
                        
                        // Si aucun OCR n'est disponible, afficher un message neutre sans mention de langue
                        if (ocrEngine == null)
                        {
                            OcrTextBox.Text = "⚠️ OCR Windows non disponible.";
                            return;
                        }
                    }
                }

                // Effectuer l'OCR avec le moteur sélectionné
                var ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);
                if (ocrResult == null || ocrResult.Lines.Count == 0)
                {
                    OcrTextBox.Text = GetLocalizedMessage("NoTextDetected");
                    return;
                }

                resultText = string.Join('\n', ocrResult.Lines.Select(l => l.Text));

                // Appliquer les corrections post-OCR et afficher
                resultText = ApplyPostOcrCorrections(resultText);
                OcrTextBox.Text = resultText;

                // Nettoyage du fichier temporaire
                try { File.Delete(processedImagePath); } catch { }
            }
            catch (Exception ex)
            {
                OcrTextBox.Text = GetLocalizedMessage("WindowsOcrError") + ex.Message;
            }
        }

        // Convertir BitmapImage -> SoftwareBitmap pour Windows OCR
        private async Task<SoftwareBitmap?> ConvertBitmapImageToSoftwareBitmap(BitmapImage bitmapImage)
        {
            try
            {
                var writeableBitmap = new WriteableBitmap(bitmapImage);

                using (var stream = new MemoryStream())
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(writeableBitmap));
                    encoder.Save(stream);
                    stream.Position = 0;

                    var randomAccessStream = new InMemoryRandomAccessStream();
                    var writer = new DataWriter(randomAccessStream);
                    writer.WriteBytes(stream.ToArray());
                    await writer.StoreAsync();
                    await writer.FlushAsync();
                    writer.DetachStream();

                    randomAccessStream.Seek(0);
                    var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(randomAccessStream);
                    return await decoder.GetSoftwareBitmapAsync();
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(System.IO.Path.Combine(tempFolder, "error_log.txt"), $"[{DateTime.Now}] Erreur conversion d'image: {ex}\n\n");
                return null;
            }
        }

        // ---------------- Post OCR corrections ----------------
        private string ApplyPostOcrCorrections(string text)
        {
            try
            {
                // Correction spécifique des séquences de "OOO" ou plus qui n'existent pas
                // Elles doivent être remplacées par "000" ou plus
                text = Regex.Replace(text, @"O{3,}", match => new string('0', match.Length));
                
                // Appliquer les corrections spécifiques pour les mots couramment mal reconnus
                // Correction de "Ted" en "Text" dans différents contextes
                text = Regex.Replace(text, @"\bTed\b", "Text", RegexOptions.IgnoreCase); // Mot isolé
                text = Regex.Replace(text, @"(?<=\b)Ted(?=box\b)", "Text", RegexOptions.IgnoreCase); // Textbox -> Tedbox
                text = Regex.Replace(text, @"(?<=\b)Ted(?=area\b)", "Text", RegexOptions.IgnoreCase); // Textarea -> Tedarea
                text = Regex.Replace(text, @"(?<=\b)Ted(?=Block\b)", "Text", RegexOptions.IgnoreCase); // TextBlock -> TedBlock
                text = Regex.Replace(text, @"(?<=\b)Ted(?=field\b)", "Text", RegexOptions.IgnoreCase); // TextField -> Tedfield
                text = Regex.Replace(text, @"(?<=\b)Ted(?=view\b)", "Text", RegexOptions.IgnoreCase); // TextView -> Tedview
                text = Regex.Replace(text, @"(?<=\b)Ted(?=file\b)", "Text", RegexOptions.IgnoreCase); // Textfile -> Tedfile
                text = Regex.Replace(text, @"(?<=\b)Ted(?=edit\b)", "Text", RegexOptions.IgnoreCase); // Textedit -> Tededit
                text = Regex.Replace(text, @"(?<=\b)Ted(?=input\b)", "Text", RegexOptions.IgnoreCase); // Textinput -> Tedinput
                text = Regex.Replace(text, @"(?<=Rich)Ted(?=\b)", "Text", RegexOptions.IgnoreCase); // RichText -> RichTed
                text = Regex.Replace(text, @"(?<=\bplain)Ted(?=\b)", "Text", RegexOptions.IgnoreCase); // plainText -> plainTed
                
                // Une approche plus généralisée pour détecter d'autres cas possibles 
                // où "Ted" apparaît dans un contexte où "Text" est plus probable
                text = Regex.Replace(text, @"(?<=\b[A-Za-z]{3,})Ted(?=\b)", match => {
                    // Si "Ted" est précédé d'un mot d'au moins 3 lettres, c'est probablement "Text"
                    return "Text";
                }, RegexOptions.IgnoreCase);
                
                // Appliquer la correction spécifique pour "LiquidF0rm" directement
                text = Regex.Replace(text, @"\bLiquidF[O0o]rm\b", "LiquidF0rm", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, @"\bLiquidForm\b", "LiquidF0rm", RegexOptions.IgnoreCase);
                
                // Correction du problème de confusion "i" vs "1"
                // Correction des cas où le chiffre "1" est reconnu comme la lettre "i"

                // 1. Remplacer "i" par "1" quand il est entouré de chiffres
                text = Regex.Replace(text, @"(?<=\d)[iIl](?=\d)", "1"); // i entre chiffres devient 1
                
                // 2. Corriger les modèles numériques courants
                text = Regex.Replace(text, @"\b[iIl](?=\d+\b)", "1"); // i suivi de chiffres
                text = Regex.Replace(text, @"(?<=\b\d+)[iIl]\b", "1"); // i à la fin d'une séquence de chiffres
                
                // 3. Corriger des modèles spécifiques où i est probablement un 1
                text = Regex.Replace(text, @"\b[iIl]{2,}\b", m => new string('1', m.Length)); // séquence de i isolée (ii, iii) devient des 1
                text = Regex.Replace(text, @"\b20[iIl]\d\b", m => "201" + m.Value.Substring(3)); // Format d'année (ex: 20i9 -> 2019)
                text = Regex.Replace(text, @"\b[iIl]9\d\d\b", m => "1" + m.Value.Substring(1)); // Format d'année (ex: i999 -> 1999)
                
                // 4. Corriger des formules ou motifs techniques (ex: IPv4, versions de logiciels)
                text = Regex.Replace(text, @"(?<!\w)([0-9]+)\.([iIl])\.([0-9]+)(?!\w)", "$1.1.$3"); // Format x.i.x -> x.1.x
                text = Regex.Replace(text, @"(?<!\w)v([iIl])\.([0-9])", "v1.$2"); // Format vi.x -> v1.x
                
                // Correction des erreurs communes de l'OCR
                // Un "O" majuscule n'est jamais isolé, donc toujours le remplacer par "0"
                // et un "0" est toujours précédé ou suivi d'un chiffre
                text = Regex.Replace(text, @"\bO\b", "0"); // Remplace "O" isolé par "0"
                text = Regex.Replace(text, @"(?<!\d)O(?!\d)", "0"); // Remplace "O" non entouré de chiffres par "0"
                text = Regex.Replace(text, @"(\d)O", "$10"); // Remplacer O par 0 après un chiffre
                text = Regex.Replace(text, @"O(\d)", "0$1"); // Remplacer O par 0 avant un chiffre
                
                // IMPORTANT: Protéger "LiquidF0rm" une nouvelle fois avant les règles qui peuvent le modifier
                text = Regex.Replace(text, @"\bLiquidF[O0o]rm\b", "LiquidF0rm", RegexOptions.IgnoreCase);
                
                // On garde les règles de conversion "0" vers "O" dans les mots, mais seulement si pas entouré de chiffres
                text = Regex.Replace(text, @"(?<![0-9])0(?![0-9])(?=[A-Za-z])", "O"); // "0" suivi d'une lettre devient "O"
                text = Regex.Replace(text, @"(?<=[A-Za-z])(?<![0-9])0(?![0-9])", "O"); // "0" précédé d'une lettre devient "O"
                text = Regex.Replace(text, @"\b([A-Za-z]+)0([A-ZaZ]+)\b(?![0-9])", "$1O$2"); // "0" au milieu d'un mot devient "O"
                
                // Appliquer la correction une troisième fois pour s'assurer qu'elle est respectée
                text = Regex.Replace(text, @"\bLiquidF[O0o]rm\b", "LiquidF0rm", RegexOptions.IgnoreCase);
                
                // Corrections pour les symboles monétaires erronément détectés dans les nombres
                text = Regex.Replace(text, @"(\d+)\$(\d+)", "$1$2"); // Supprime le $ entre des chiffres (ex: 3$10 -> 310)
                text = Regex.Replace(text, @"(\d+)€(\d+)", "$1$2");  // Supprime le € entre des chiffres
                text = Regex.Replace(text, @"(\d+)£(\d+)", "$1$2");  // Supprime le £ entre des chiffres
                text = Regex.Replace(text, @"(\d+)¥(\d+)", "$1$2");  // Supprime le ¥ entre des chiffres
                
                // Corrections spécifiques pour le problème C# reconnu comme (#
                text = text.Replace("(€", "C#");
                text = text.Replace("(E", "C#");
                text = text.Replace("(£", "C#");
                text = text.Replace("(#", "C#");
                text = text.Replace("<€", "C#");
                text = text.Replace("<#", "C#");
                text = text.Replace("C€", "C#");
                text = text.Replace("C €", "C#");
                text = text.Replace("c#", "C#");
                text = text.Replace("cé", "C#");
                text = text.Replace("Cë", "C#");

                // Recherche de patterns supplémentaires pour C#
                var csharpPattern = new Regex(@"\bC[\s\.\-_]*(\(|<|{|\[|€|£|\$|#|\+|\*)\b");
                text = csharpPattern.Replace(text, "C#");

                // Corrections avancées basées sur des expressions régulières
                text = Regex.Replace(text, @"l(\d)", "1$1"); // Remplacer l par 1 avant un chiffre
                
                // Corrections supplémentaires pour "l" reconnu comme "1"
                text = Regex.Replace(text, @"[iI](?=\d\b)", "1"); // i suivi d'un seul chiffre à la fin d'un mot
                text = Regex.Replace(text, @"(?<=\b\d)[iI](?=\d)", "1"); // i au milieu d'une séquence de chiffres
                
                // Vérification pour éviter de transformer des mots réels avec "i" en "1"
                // Liste de mots à protéger adaptée à la langue d'interface
                string[] wordsToProtect;
                if (isEnglish)
                {
                    // Mots anglais courants contenant "i"
                    wordsToProtect = new[] { 
                        "in", "it", "is", "if", "with", "this", "will", "which", "time", "like", 
                        "think", "did", "first", "into", "their", "him", "his", "itself", "simple",
                        "give", "might", "list", "bit", "win", "sit", "big", "fit", "hit", "mix", 
                        "fix", "fill", "still", "while", "skin"
                    };
                }
                else
                {
                    // Mots français courants contenant "i"
                    wordsToProtect = new[] { 
                        "si", "ni", "ici", "qui", "aussi", "ainsi", "midi", "parmi", "quasi", 
                        "mini", "maxi", "kiwi", "wiki", "ski", "alibi", "anti", "api", "bit", 
                        "prix", "vis", "dit", "riz", "dix", "fil", "ville", "fille"
                    };
                }
                
                // Restaurer les mots protégés (ne s'applique qu'aux mots entiers)
                foreach (var word in wordsToProtect)
                {
                    // Pattern qui recherche le mot avec un "1" à la place du "i" (mot entier uniquement)
                    string pattern = @"\b" + Regex.Escape(word.Replace("i", "1")) + @"\b";
                    text = Regex.Replace(text, pattern, word, RegexOptions.IgnoreCase);
                }

                // Rechercher les motifs de programmation courants pour les corriger
                text = Regex.Replace(text, @"\bC\s+#\b", "C#"); // Corriger "C #" en "C#"
                text = Regex.Replace(text, @"\b[cC]\+\+?\b", "C++"); // Normaliser C++
                text = Regex.Replace(text, @"\b[jJ][aA][vV][aA]\s*[sS][cC][rR][iI][pP][tT]\b", "JavaScript"); // Normaliser JavaScript

                // Corrections pour d'autres symboles spéciaux
                text = text.Replace("＃", "#"); // Unicode alternatif vers ASCII standard
                text = text.Replace("＠", "@"); // Unicode alternatif vers ASCII standard
                text = text.Replace("＆", "&"); // Unicode alternatif vers ASCII standard
                text = text.Replace("％", "%"); // Unicode alternatif vers ASCII standard

                // Vérification finale pour LiquidF0rm
                text = Regex.Replace(text, @"\bLiquidF[O0o]rm\b", "LiquidF0rm", RegexOptions.IgnoreCase);

                // Nettoyer les espaces multiples mais PAS les sauts de ligne
                text = Regex.Replace(text, @"[ \t]+", " ");

                // Nettoyer les retours à la ligne Windows en retours à la ligne Unix
                text = text.Replace("\r\n", "\n");

                // Limiter seulement les sauts de ligne excessifs (plus de 3) à 3 sauts de ligne
                text = Regex.Replace(text, @"\n{4,}", "\n\n\n");

                return text;
            }
            catch (Exception ex)
            {
                // Log l'erreur mais retourne le texte original pour ne pas bloquer le traitement
                File.AppendAllText(Path.Combine(tempFolder, "ocr_correction_error_log.txt"),
                    $"[{DateTime.Now}] Erreur correction OCR: {ex.Message}\n{ex.StackTrace}\n\n");
                return text;
            }
        }

        // ---------------- Utilitaires UI & Divers ----------------
        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (OcrTextBox != null)
            {
                if (!string.IsNullOrEmpty(OcrTextBox.SelectedText))
                {
                    // Copier le texte sélectionné
                    Clipboard.SetText(OcrTextBox.SelectedText);
                }
                else if (!string.IsNullOrEmpty(OcrTextBox.Text))
                {
                    // Copier tout le texte si rien n'est sélectionné
                    Clipboard.SetText(OcrTextBox.Text);
                }
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            ImagePreview.Source = null;
            OcrTextBox.Text = string.Empty;
        }

        private Border? GetBorderRight() => this.FindName("BorderRight") as Border;
        private Grid? GetGridContent() => this.FindName("GridContent") as Grid;

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Key == Key.C && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                e.Handled = true;
                StartScreenCapture();
            }
        }

        private void StartScreenCapture()
        {
            CaptureButton_Click(this, new RoutedEventArgs());
        }

        private void ThemeButton_Click(object sender, RoutedEventArgs e)
        {
            isDarkTheme = !isDarkTheme;
            ApplyTheme();
        }

        private void ApplyTheme()
        {
            var borderLeft = this.FindName("BorderLeft") as Border;
            var borderRight = this.FindName("BorderRight") as Border;
            var gridContent = this.FindName("GridContent") as Grid;
            var ocrTextBox = this.FindName("OcrTextBox") as TextBox;

            if (borderLeft != null && borderRight != null && gridContent != null && ocrTextBox != null)
            {
                if (isDarkTheme)
                {
                    borderLeft.Background = new SolidColorBrush(System.Windows.Media.Colors.Black);
                    borderRight.Background = new SolidColorBrush(System.Windows.Media.Colors.Black);
                    ocrTextBox.Background = new SolidColorBrush(System.Windows.Media.Colors.Black);
                    ocrTextBox.Foreground = new SolidColorBrush(System.Windows.Media.Colors.White);
                    ocrTextBox.BorderBrush = new SolidColorBrush(System.Windows.Media.Colors.DimGray);

                    gridContent.Background = new LinearGradientBrush
                    {
                        StartPoint = new System.Windows.Point(0.5, 0),
                        EndPoint = new System.Windows.Point(0.5, 1),
                        GradientStops = new GradientStopCollection
                        {
                            new GradientStop(System.Windows.Media.Colors.Gray, 0),
                            new GradientStop(System.Windows.Media.Colors.Black, 1)
                        }
                    };
                }
                else
                {
                    borderLeft.Background = new SolidColorBrush(System.Windows.Media.Colors.WhiteSmoke);
                    borderRight.Background = new SolidColorBrush(System.Windows.Media.Colors.White);
                    ocrTextBox.Background = new SolidColorBrush(System.Windows.Media.Colors.White);
                    ocrTextBox.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Black);
                    ocrTextBox.BorderBrush = new SolidColorBrush(System.Windows.Media.Colors.DarkGray);

                    gridContent.Background = new LinearGradientBrush
                    {
                        StartPoint = new System.Windows.Point(0.5, 0),
                        EndPoint = new System.Windows.Point(0.5, 1),
                        GradientStops = new GradientStopCollection
                        {
                            new GradientStop(System.Windows.Media.Color.FromRgb(158, 106, 255), 0),
                            new GradientStop(System.Windows.Media.Color.FromRgb(141, 230, 253), 1)
                        }
                    };
                }
            }
            else
            {
                // don't spam MessageBox in production; leave a silent fallback
            }
        }

        private void TestImageManipulation(string imagePath)
        {
            try
            {
                using (var image = new MagickImage(imagePath))
                {
                    image.Blur(0, 1.5);
                    using (var ms = new MemoryStream(image.ToByteArray()))
                    {
                        var bmp = new Bitmap(ms);
                        // MonImageControl.Source = BitmapToImageSource(bmp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors du traitement de l'image : " + ex.Message);
            }
        }

        private BitmapSource BitmapToImageSource(Bitmap bmp)
        {
            IntPtr hBitmap = bmp.GetHbitmap();
            try
            {
                return Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private async void CaptureButton_Click(object sender, RoutedEventArgs e)
        {
            var overlay = new OverlayWindow();
            overlay.CaptureCompleted += bitmap =>
            {
                try
                {
                    // Supprimer les fichiers d'autosauvegarde existants
                    DeleteAutosaveSession();
                    
                    string tempPath = System.IO.Path.Combine(tempFolder, "temp_capture_" + Guid.NewGuid().ToString() + ".png");
                    bitmap.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);

                    Dispatcher.Invoke(async () =>
                    {
                        ImagePreview.Source = BitmapToImageSource(bitmap);
                        try
                        {
                            await RunWindowsOcr(tempPath);
                        }
                        catch (Exception ocrEx)
                        {
                            OcrTextBox.Text = "⚠️ L'OCR a échoué sur la capture d'écran: " + ocrEx.Message;
                        }
                    });
                }
                catch (Exception captureEx)
                {
                    Dispatcher.Invoke(() => MessageBox.Show("Erreur lors du traitement de la capture: " + captureEx.Message));
                }
            };
            overlay.ShowDialog();
        }

        private void Language_Click(object sender, RoutedEventArgs e)
        {
            // Basculer entre français et anglais
            isEnglish = !isEnglish;

            // Mettre à jour le texte du bouton de langue
            LanguageButton.Content = isEnglish ? "English" : "Français";

            // Appliquer les changements de langue
            ApplyLanguage();

            // Sauvegarder la préférence de langue
            Properties.Settings.Default.LanguageParameter = isEnglish ? "en" : "fr";
            Properties.Settings.Default.Save();
        }

        private void ApplyLanguage()
        {
            if (isEnglish)
            {
                // Appliquer les textos en anglais
                BrowseButton.Content = "Browse...";
                CaptureButton.Content = "Screenshot";
                SaveasButton.Content = "Save as...";
                SaveTextButton.Content = "Save text as..."; // Texte du bouton d'enregistrement de texte en anglais
                CopyButton.Content = "Copy";
                ThemeButton.Content = "Theme";
                SaveButton.Content = "Save";  // Traduction du bouton SaveButton
            }
            else
            {
                // Appliquer les textos en français
                BrowseButton.Content = "Parcourir...";
                CaptureButton.Content = "Capture d'écran";
                SaveasButton.Content = "Enregistrer sous...";
                SaveTextButton.Content = "Enregistrer en .txt"; // Texte du bouton d'enregistrement de texte en français
                CopyButton.Content = "Copier";
                ThemeButton.Content = "Thème";
                SaveButton.Content = "Sauvegarder";  // Traduction du bouton SaveButton
            }
        }

        private void CleanupOldTempFiles()
        {
            try
            {
                if (Directory.Exists(tempFolder))
                {
                    var files = Directory.GetFiles(tempFolder);
                    foreach (var file in files)
                    {
                        try
                        {
                            var info = new FileInfo(file);
                            // Supprime les fichiers de plus de 2 jours
                            if (info.LastWriteTime < DateTime.Now.AddDays(-2))
                                File.Delete(file);
                        }
                        catch { /* Ignorer les erreurs de suppression individuelles */ }
                    }
                }
            }
            catch { /* Ignorer les erreurs globales */ }
        }

        // Ajoutez cette méthode utilitaire dans la classe MainWindow (par exemple juste avant le constructeur)
        private double ExtractNumericValue(string? value, double defaultValue)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;
            if (double.TryParse(value, out double result))
                return result;
            return defaultValue;
        }

        /// <summary>
        /// Détecte si une image MagickImage a majoritairement un fond sombre.
        /// Technique : encode l'image en PNG en mémoire, charge dans SKBitmap et échantillonne.
        /// Retourne true si la luminance moyenne < threshold (ajustable).
        /// </summary>
        private bool IsDarkBackground(MagickImage magickImg, double threshold = 0.40)
        {
            try
            {
                // Encode en PNG en mémoire (safe)
                using (var ms = new MemoryStream())
                {
                    magickImg.Write(ms, MagickFormat.Png);
                    ms.Position = 0;

                    using (var sk = SkiaSharp.SKBitmap.Decode(ms))
                    {
                        if (sk == null || sk.Width == 0 || sk.Height == 0)
                            return false;

                        int stepX = Math.Max(1, sk.Width / 20);
                        int stepY = Math.Max(1, sk.Height / 20);
                        double sum = 0;
                        int count = 0;

                        for (int y = 0; y < sk.Height; y += stepY)
                        {
                            for (int x = 0; x < sk.Width; x += stepX)
                            {
                                var c = sk.GetPixel(x, y);
                                // SKColor channels are 0..255
                                double lum = 0.2126 * c.Red + 0.7152 * c.Green + 0.0722 * c.Blue;
                                sum += lum;
                                count++;
                            }
                        }

                        if (count == 0) return false;
                        double mean = (sum / count) / 255.0; // normalize 0..1
                        return mean < threshold;
                    }
                }
            }
            catch
            {
                // best effort — si erreur, ne pas forcer inversion
                return false;
            }
        }

        // Variante douce (préserve les détails, pas d'inversion agressive)
        private void PreprocessImageForOcr_Douce(string originalImagePath, string processedImagePath)
        {
            using (var mag = new MagickImage(originalImagePath))
            {
                mag.Density = new Density(300, 300);

                // Resize léger si petit
                if (Math.Max(mag.Width, mag.Height) < 900)
                {
                    uint newW = (uint)Math.Min(mag.Width * 2, 4000);
                    uint newH = (uint)Math.Round((double)mag.Height * newW / mag.Width);
                    mag.FilterType = FilterType.Cubic;
                    mag.Resize(newW, newH);
                }

                // Conserver la couleur pour OCR
                // Netteté douce
                mag.UnsharpMask(0.3, 0.5, 0.5, 0.02);

                mag.Format = MagickFormat.Png;
                mag.Write(processedImagePath);
            }
        }

        // Variante moyenne (équilibre : inversion si fond sombre, contraste modéré, resize x2)
        private void PreprocessImageForOcr_Moyenne(string originalImagePath, string processedImagePath)
        {
            using (var mag = new MagickImage(originalImagePath))
            {
                mag.Density = new Density(300, 300);

                if (Math.Max(mag.Width, mag.Height) < 900)
                {
                    uint newW = (uint)Math.Min(mag.Width * 2, 5000);
                    uint newH = (uint)Math.Round((double)mag.Height * newW / mag.Width);
                    mag.FilterType = FilterType.Lanczos;
                    mag.Resize(newW, newH);
                }

                // Détection fond sombre
                bool isDarkBg = IsDarkBackground(mag, 0.35);
                if (isDarkBg)
                {
                    mag.Negate();
                    mag.SigmoidalContrast(2.0, 0.5, Channels.All);
                }

                mag.ReduceNoise(1);
                mag.ColorType = ColorType.Grayscale;
                mag.UnsharpMask(0.5, 0.5, 0.7, 0.02);

                mag.Format = MagickFormat.Png;
                mag.Write(processedImagePath);
            }
        }

        // Variante agressive (fort contraste, seuillage adaptatif, pour images sombres/brouillées)
        private void PreprocessImageForOcr_Agressive(string originalImagePath, string processedImagePath)
        {
            using (var mag = new MagickImage(originalImagePath))
            {
                mag.Density = new Density(300, 300);

                if (Math.Max(mag.Width, mag.Height) < 800)
                {
                    uint newW = (uint)Math.Min(mag.Width * 3, 5000);
                    uint newH = (uint)Math.Round((double)mag.Height * newW / mag.Width);
                    mag.FilterType = FilterType.Lanczos;
                    mag.Resize(newW, newH);
                }

                // Inversion si nécessaire
                bool isDarkBg = IsDarkBackground(mag, 0.45);
                if (isDarkBg)
                {
                    mag.Negate();
                    mag.SigmoidalContrast(3.0, 0.5, Channels.All);
                }

                // Seuillage adaptatif agressif
                mag.AdaptiveThreshold(15, 15, 10);

                // Grayscale obligatoire ici
                mag.ColorType = ColorType.Grayscale;

                mag.Format = MagickFormat.Png;
                mag.Write(processedImagePath);
            }
        }

        private void Saveas_Click(object sender, RoutedEventArgs e)
        {
            // Vérifier si une image est chargée
            if (ImagePreview.Source == null)
            {
                MessageBox.Show(GetLocalizedMessage("NoImageToSave"),
                               isEnglish ? "Information" : "Information",
                               MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Créer une boîte de dialogue pour choisir l'emplacement et le format
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Title = GetLocalizedMessage("SaveImageTitle"),
                    Filter = isEnglish
                           ? "JPEG Image (*.jpg)|*.jpg|PNG Image (*.png)|*.png"
                           : "Image JPEG (*.jpg)|*.jpg|Image PNG (*.png)|*.png",
                    DefaultExt = ".jpg",
                    AddExtension = true
                };

                // Si l'utilisateur clique sur OK
                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    string extension = System.IO.Path.GetExtension(filePath).ToLowerInvariant();

                    // Obtenir l'image à partir de la source
                    BitmapSource? bitmapSource = ImagePreview.Source as BitmapSource;
                    if (bitmapSource != null)
                    {
                        System.Windows.Media.Imaging.BitmapEncoder encoder;
                        if (extension == ".png")
                        {
                            encoder = new PngBitmapEncoder();
                        }
                        else // .jpg par défaut
                        {
                            JpegBitmapEncoder jpegEncoder = new JpegBitmapEncoder();
                            jpegEncoder.QualityLevel = 90; // Qualité JPEG (0-100)
                            encoder = jpegEncoder;
                        }

                        encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapSource));

                        // Enregistrer l'image dans le fichier
                        using (var fileStream = new FileStream(filePath, FileMode.Create))
                        {
                            encoder.Save(fileStream);
                        }

                        MessageBox.Show(GetLocalizedMessage("SaveImageSuccess") + filePath,
                                       isEnglish ? "Success" : "Succès",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetLocalizedMessage("SaveImageError") + ex.Message,
                               isEnglish ? "Error" : "Erreur",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Méthode du bouton de sauvegarde
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Vérifier s'il y a une image à sauvegarder
            if (ImagePreview.Source == null)
            {
                MessageBox.Show(
                    isEnglish ? "No image to save." : "Aucune image à sauvegarder.",
                    isEnglish ? "Information" : "Information",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Sauvegarder l'image
                BitmapSource? bitmapSource = ImagePreview.Source as BitmapSource;
                if (bitmapSource != null)
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    // Utilisez le namespace complet pour éviter l'ambiguïté
                    encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapSource));

                    using (var fileStream = new FileStream(autosaveImagePath, FileMode.Create))
                    {
                        encoder.Save(fileStream);
                    }
                }

                // Sauvegarder le texte
                File.WriteAllText(autosaveTextPath, OcrTextBox.Text);

                // Sauvegarder d'autres informations de configuration si nécessaire
                File.WriteAllText(autosaveConfigPath, 
                    $"SaveTime={DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                    $"Theme={isDarkTheme}\n" +
                    $"Language={isEnglish}");

                // Afficher un message de confirmation
                MessageBox.Show(
                    isEnglish ? "Session saved successfully." : "Session sauvegardée avec succès.",
                    isEnglish ? "Success" : "Succès",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    (isEnglish ? "Error saving session: " : "Erreur lors de la sauvegarde de la session : ") + ex.Message,
                    isEnglish ? "Error" : "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveTextButton_Click(object sender, RoutedEventArgs e)
        {
            // Vérifier si du texte est présent
            if (string.IsNullOrWhiteSpace(OcrTextBox.Text))
            {
                MessageBox.Show(
                    isEnglish ? "No text to save." : "Aucun texte à enregistrer.",
                    isEnglish ? "Information" : "Information",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Créer une boîte de dialogue pour choisir l'emplacement du fichier
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Title = isEnglish ? "Save text as" : "Enregistrer le texte sous",
                    Filter = isEnglish ? "Text files (*.txt)|*.txt|All files (*.*)|*.*" : "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*",
                    DefaultExt = ".txt",
                    AddExtension = true
                };

                // Si l'utilisateur clique sur OK
                if (saveFileDialog.ShowDialog() == true)
                {
                    // Enregistrer le texte dans le fichier
                    File.WriteAllText(saveFileDialog.FileName, OcrTextBox.Text);

                    // Afficher un message de confirmation
                    MessageBox.Show(
                        (isEnglish ? "Text successfully saved at: " : "Texte enregistré avec succès à : ") + saveFileDialog.FileName,
                        isEnglish ? "Success" : "Succès",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                // Afficher un message d'erreur en cas de problème
                MessageBox.Show(
                    (isEnglish ? "Error saving text: " : "Erreur lors de l'enregistrement du texte : ") + ex.Message,
                    isEnglish ? "Error" : "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 1. Ajoutez cette méthode manquante pour restaurer la dernière session
        private void RestoreLastSession()
        {
            try
            {
                // Vérifier si des fichiers d'autosauvegarde existent
                if (File.Exists(autosaveImagePath) && File.Exists(autosaveTextPath))
                {
                    // Restaurer l'image
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.UriSource = new Uri(autosaveImagePath);
                    bitmap.EndInit();
                    bitmap.Freeze();
                    ImagePreview.Source = bitmap;

                    // Restaurer le texte
                    OcrTextBox.Text = File.ReadAllText(autosaveTextPath);
                }
            }
            catch (Exception ex)
            {
                // En cas d'erreur, nous continuons silencieusement - pas besoin d'alarmer l'utilisateur
                File.AppendAllText(Path.Combine(tempFolder, "restore_error_log.txt"),
                    $"[{DateTime.Now}] Erreur restauration session: {ex.Message}\n{ex.StackTrace}\n\n");
            }
        }

        // Dictionnaires de messages localisés
        private readonly Dictionary<string, string> frenchMessages = new Dictionary<string, string>
        {
            // Messages d'erreur OCR
            ["NoTextDetected"] = "Aucun texte détecté.",
            ["PreprocessedFileNotFound"] = "⚠️ Le fichier prétraité n'existe pas.",
            ["ImageConversionFailed"] = "⚠️ Conversion image pour OCR Windows échouée.",
            ["WindowsOcrNotAvailable"] = "⚠️ OCR Windows (fr-FR) non disponible.",
            ["OcrLanguagePackInstruction"] = "Pour utiliser l'OCR en anglais, assurez-vous que le pack de langue anglais est installé sur votre système Windows. Vous pouvez l'installer via Paramètres > Heure et langue > Langue > Ajouter une langue.",
            ["OcrFailedOnScreenshot"] = "⚠️ OCR échoué sur la capture d'écran: ",
            ["OcrExecutionError"] = "Erreur lors de l'exécution de l'OCR : ",
            ["WindowsOcrError"] = "⚠️ Erreur OCR Windows : ",

            // Messages pour l'enregistrement d'image
            ["NoImageToSave"] = "Aucune image à enregistrer.",
            ["SaveImageTitle"] = "Enregistrer l'image sous",
            ["SaveImageSuccess"] = "Image enregistrée avec succès à :\n",
            ["SaveImageError"] = "Erreur lors de l'enregistrement de l'image : ",

            // Messages d'erreur divers
            ["UnsupportedFileFormat"] = "Format de fichier non pris en charge.",
            ["ImageLoadError"] = "Une erreur est survenue lors du chargement de l'image : ",
            ["ImageProcessingError"] = "Erreur lors du traitement de l'image : ",
            ["ScreenCaptureError"] = "Erreur lors du traitement de la capture: "
        };

        private readonly Dictionary<string, string> englishMessages = new Dictionary<string, string>
        {
            // OCR error messages
            ["NoTextDetected"] = "No text detected.",
            ["PreprocessedFileNotFound"] = "⚠️ Preprocessed file does not exist.",
            ["ImageConversionFailed"] = "⚠️ Image conversion for Windows OCR failed.",
            ["WindowsOcrNotAvailable"] = "⚠️ Windows OCR (en-US) not available.",
            ["OcrLanguagePackInstruction"] = "To use English OCR, make sure the English language pack is installed on your Windows system. You can install it via Settings > Time & Language > Language > Add a language.",
            ["OcrFailedOnScreenshot"] = "⚠️ OCR failed on screenshot: ",
            ["OcrExecutionError"] = "Error during OCR execution: ",
            ["WindowsOcrError"] = "⚠️ Windows OCR error: ",
            
            // Image saving messages
            ["NoImageToSave"] = "No image to save.",
            ["SaveImageTitle"] = "Save image as",
            ["SaveImageSuccess"] = "Image successfully saved at:\n",
            ["SaveImageError"] = "Error saving image: ",

            // Miscellaneous error messages
            ["UnsupportedFileFormat"] = "Unsupported file format.",
            ["ImageLoadError"] = "An error occurred while loading the image: ",
            ["ImageProcessingError"] = "Error processing image: ",
            ["ScreenCaptureError"] = "Error processing screen capture: ",

            // Messages de sauvegarde de texte
            ["SaveTextSuccess"] = "Text successfully saved at:\n",
            ["SaveTextError"] = "Error saving text: ",
        };

        // Méthode d'assistance pour obtenir un message dans la langue sélectionnée
        private string GetLocalizedMessage(string messageKey)
        {
            return isEnglish
                ? englishMessages.TryGetValue(messageKey, out string? englishMsg) ? englishMsg : messageKey
                : frenchMessages.TryGetValue(messageKey, out string? frenchMsg) ? frenchMsg : messageKey;
        }
    }
}

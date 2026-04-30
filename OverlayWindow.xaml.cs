using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Runtime.Versioning;

// Pour éviter les ambiguïtés
using SWM = System.Windows.Media;
using SWVisibility = System.Windows.Visibility;
using SWMColors = System.Windows.Media.Colors;
using SWControls = System.Windows.Controls;
using SWDuration = System.Windows.Duration;

namespace WpfImageToText
{
    [SupportedOSPlatform("windows")]
    public partial class OverlayWindow : System.Windows.Window
    {
        private System.Windows.Point startPoint;
        private System.Windows.Shapes.Rectangle selectionRect;
        private Border? dimensionsDisplay;
        private TextBlock? dimensionsText;

        // Modifiez la déclaration de l'événement pour accepter null
        public event Action<System.Drawing.Bitmap>? CaptureCompleted;

        public OverlayWindow()
        {
            InitializeComponent();

            this.KeyDown += OverlayWindow_KeyDown;
            this.MouseLeftButtonDown += OverlayWindow_MouseLeftButtonDown;
            this.MouseMove += OverlayWindow_MouseMove;
            this.MouseLeftButtonUp += OverlayWindow_MouseLeftButtonUp;

            // Personnalisation du rectangle de sélection
            selectionRect = new System.Windows.Shapes.Rectangle
            {
                Stroke = new SolidColorBrush(SWMColors.Indigo),
                StrokeThickness = 1,
                Visibility = SWVisibility.Hidden,
                Fill = new SolidColorBrush(SWM.Color.FromArgb(40, 0, 120, 215)) // Remplissage semi-transparent
            };
            
            // Ajouter un effet d'ombre au rectangle
            selectionRect.Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                Color = SWMColors.Black,
                Direction = 315,
                ShadowDepth = 5,
                Opacity = 0.3,
                BlurRadius = 5
            };
            
            SelectionCanvas.Children.Add(selectionRect);
            
            // Récupérer la référence aux éléments d'affichage des dimensions
            dimensionsDisplay = this.FindName("DimensionsDisplay") as Border;
            dimensionsText = this.FindName("DimensionsText") as TextBlock;

            this.Cursor = Cursors.Cross; // Utilise un curseur en croix pour la sélection

            // Appliquer l'apparence moderne
            ApplyModernAppearance();
        }

        private void ApplyModernAppearance()
        {
            try
            {
                // Utiliser une couleur par défaut à la place
                var wpfAccentColor = SWM.Color.FromRgb(0, 120, 215); // Couleur bleue similaire à l'accent Windows
                
                // Exemple d'utilisation de cette couleur (si nécessaire)
                // selectionRect.Stroke = new SolidColorBrush(wpfAccentColor);
            }
            catch
            {
                // Fallback en cas d'erreur
            }
        }

        private void OverlayWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            startPoint = e.GetPosition(SelectionCanvas);
            SWControls.Canvas.SetLeft(selectionRect, startPoint.X);
            SWControls.Canvas.SetTop(selectionRect, startPoint.Y);
            selectionRect.Width = 0;
            selectionRect.Height = 0;
            selectionRect.Visibility = SWVisibility.Visible;
            this.CaptureMouse();
            
            // Masquer l'affichage des dimensions au début de la sélection
            if (dimensionsDisplay != null)
                dimensionsDisplay.Visibility = SWVisibility.Collapsed;
        }

        private void OverlayWindow_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var pos = e.GetPosition(SelectionCanvas);
                double x = Math.Min(pos.X, startPoint.X);
                double y = Math.Min(pos.Y, startPoint.Y);
                double w = Math.Abs(pos.X - startPoint.X);
                double h = Math.Abs(pos.Y - startPoint.Y);

                SWControls.Canvas.SetLeft(selectionRect, x);
                SWControls.Canvas.SetTop(selectionRect, y);
                selectionRect.Width = w;
                selectionRect.Height = h;
                
                // Mettre à jour et afficher les dimensions
                if (dimensionsDisplay != null && dimensionsText != null)
                {
                    dimensionsText.Text = $"{(int)w} × {(int)h} px";
                    dimensionsDisplay.Visibility = SWVisibility.Visible;
                }
            }
        }

        private void OverlayWindow_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            this.ReleaseMouseCapture();
            selectionRect.Visibility = SWVisibility.Hidden;
            
            // Masquer l'affichage des dimensions
            if (dimensionsDisplay != null)
                dimensionsDisplay.Visibility = SWVisibility.Collapsed;

            var endPoint = e.GetPosition(SelectionCanvas);
            int x = (int)Math.Min(startPoint.X, endPoint.X);
            int y = (int)Math.Min(startPoint.Y, endPoint.Y);
            int width = (int)Math.Abs(endPoint.X - startPoint.X);
            int height = (int)Math.Abs(endPoint.Y - startPoint.Y);

            if (width < 10 || height < 10) // Modification : taille minimale requise
            {
                // Afficher un message d'information si la sélection est trop petite
                if (width > 0 && height > 0)
                {
                    MessageBox.Show("La zone sélectionnée est trop petite. Veuillez sélectionner une zone plus grande.", 
                                    "Sélection trop petite", 
                                    MessageBoxButton.OK, 
                                    MessageBoxImage.Information);
                }
                this.Close();
                return;
            }

            // Animation simple avant la capture (optionnelle)
            System.Windows.Media.Animation.DoubleAnimation fadeAnimation = new System.Windows.Media.Animation.DoubleAnimation
            {
                From = 0.8,
                To = 0,
                Duration = new SWDuration(TimeSpan.FromMilliseconds(300))
            };
            
            // Remplacer la section de capture d'écran existante par celle-ci
            fadeAnimation.Completed += (s, args) =>
            {
                CaptureScreenshot(x, y, width, height);
                this.Close();
            };
            
            this.Background.BeginAnimation(SolidColorBrush.OpacityProperty, fadeAnimation);
        }

        [SupportedOSPlatform("windows6.1")]
        private void CaptureScreenshot(int x, int y, int width, int height)
        {
            try
            {
                // Capture d'écran avec une meilleure qualité
                using (var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        // Améliorer la qualité de la capture
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        
                        // Capturer l'écran
                        g.CopyFromScreen((int)this.Left + x, (int)this.Top + y, 0, 0, 
                            new System.Drawing.Size(width, height), 
                            CopyPixelOperation.SourceCopy);
                    }

                    // Utiliser l'image directement sans redimensionnement
                    CaptureCompleted?.Invoke((System.Drawing.Bitmap)bmp.Clone());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la capture d'écran : " + ex.Message, 
                                "Erreur de capture", 
                                MessageBoxButton.OK, 
                                MessageBoxImage.Error);
            }
        }

        private void OverlayWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }
    }
}

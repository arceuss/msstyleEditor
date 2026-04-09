using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace msstyleEditor
{
	class ImageControl : Panel
	{
		public enum BackgroundStyle
		{
			Color,
			Chessboard
		}

        private const float ZoomStep = 0.1f;

        private readonly Size m_cellSize = new Size(16, 16);
        private BackgroundStyle m_background = BackgroundStyle.Chessboard;
        private Rectangle? m_highlightArea;

        public event EventHandler ZoomChanged;

        public BackgroundStyle Background
        {
            get
            {
                return m_background;
            }
            set
            {
                if (m_background == value)
                {
                    return;
                }

                m_background = value;
                Invalidate();
            }
        }

        public bool EnableInteractiveZoom { get; set; }

        public float MinZoom { get; set; } = 0.5f;
        public float MaxZoom { get; set; } = 8.0f;
		
        private float m_zoomFactor = 1.0f;
        public float ZoomFactor
        {
            get
            {
                return m_zoomFactor;
            }
            set
            {
                SetZoomFactor(value);
            }
        }

        public Rectangle? HighlightArea
        {
            get
            {
                return m_highlightArea;
            }
            set
            {
                if (m_highlightArea == value)
                {
                    return;
                }

                m_highlightArea = value;
                Invalidate();
            }
        }


		public ImageControl()
        {
            AutoScroll = true;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true);
        }


        public void ResetZoom()
        {
            SetZoomFactor(1.0f, null, true);
        }


        public void SetZoomFactor(float value)
        {
            SetZoomFactor(value, null, false);
        }


        private void SetZoomFactor(float value, Point? focusPoint, bool recenterViewport)
        {
            float clampedZoom = Math.Max(MinZoom, Math.Min(MaxZoom, value));
            float previousZoom = m_zoomFactor;
            bool zoomChanged = Math.Abs(previousZoom - clampedZoom) > 0.0001f;

            if (!zoomChanged && !recenterViewport)
            {
                return;
            }

            Point focus = GetFocusPoint(focusPoint);
            PointF sourcePoint = GetSourcePoint(previousZoom, focus);

            m_zoomFactor = clampedZoom;
            ApplyLayout(sourcePoint, focus, recenterViewport);

            if (zoomChanged)
            {
                ZoomChanged?.Invoke(this, EventArgs.Empty);
            }
        }


        private Size GetScaledImageSize(float zoomFactor)
        {
            if (BackgroundImage == null)
            {
                return Size.Empty;
            }

            return new Size(
                Math.Max(1, (int)Math.Round(BackgroundImage.Width * zoomFactor)),
                Math.Max(1, (int)Math.Round(BackgroundImage.Height * zoomFactor))
            );
        }


        private Size GetContentSize(float zoomFactor)
        {
            Size scaledImageSize = GetScaledImageSize(zoomFactor);
            return new Size(
                Math.Max(ClientSize.Width, scaledImageSize.Width),
                Math.Max(ClientSize.Height, scaledImageSize.Height)
            );
        }


        private Point GetImageOrigin(Size contentSize, Size scaledImageSize)
        {
            return new Point(
                Math.Max((contentSize.Width - scaledImageSize.Width) / 2, 0),
                Math.Max((contentSize.Height - scaledImageSize.Height) / 2, 0)
            );
        }


        private Rectangle GetImageBounds(float zoomFactor)
        {
            if (BackgroundImage == null)
            {
                return Rectangle.Empty;
            }

            Size scaledImageSize = GetScaledImageSize(zoomFactor);
            Size contentSize = GetContentSize(zoomFactor);
            Point imageOrigin = GetImageOrigin(contentSize, scaledImageSize);
            imageOrigin.Offset(AutoScrollPosition);

            return new Rectangle(imageOrigin, scaledImageSize);
        }


        private Rectangle ScaleRectangle(Rectangle rectangle, float zoomFactor)
        {
            return new Rectangle(
                (int)Math.Round(rectangle.X * zoomFactor),
                (int)Math.Round(rectangle.Y * zoomFactor),
                Math.Max(1, (int)Math.Round(rectangle.Width * zoomFactor)),
                Math.Max(1, (int)Math.Round(rectangle.Height * zoomFactor))
            );
        }


        private Point GetFocusPoint(Point? focusPoint)
        {
            if (!focusPoint.HasValue)
            {
                return new Point(ClientSize.Width / 2, ClientSize.Height / 2);
            }

            return new Point(
                Math.Max(0, Math.Min(ClientSize.Width, focusPoint.Value.X)),
                Math.Max(0, Math.Min(ClientSize.Height, focusPoint.Value.Y))
            );
        }


        private PointF GetSourcePoint(float zoomFactor, Point focusPoint)
        {
            if (BackgroundImage == null)
            {
                return PointF.Empty;
            }

            Rectangle imageBounds = GetImageBounds(zoomFactor);
            float x = (focusPoint.X - imageBounds.Left) / zoomFactor;
            float y = (focusPoint.Y - imageBounds.Top) / zoomFactor;

            return new PointF(
                Math.Max(0.0f, Math.Min(BackgroundImage.Width, x)),
                Math.Max(0.0f, Math.Min(BackgroundImage.Height, y))
            );
        }


        private Point ClampScrollOffset(Point scrollOffset, Size contentSize)
        {
            return new Point(
                Math.Max(0, Math.Min(scrollOffset.X, Math.Max(contentSize.Width - ClientSize.Width, 0))),
                Math.Max(0, Math.Min(scrollOffset.Y, Math.Max(contentSize.Height - ClientSize.Height, 0)))
            );
        }


        private Point GetCenteredScrollOffset(Size contentSize)
        {
            return new Point(
                Math.Max((contentSize.Width - ClientSize.Width) / 2, 0),
                Math.Max((contentSize.Height - ClientSize.Height) / 2, 0)
            );
        }


        private void UpdateLayout(bool recenterViewport)
        {
            Size contentSize = BackgroundImage != null ? GetContentSize(m_zoomFactor) : ClientSize;
            Point scrollOffset = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y);

            AutoScrollMinSize = contentSize;

            if (BackgroundImage == null)
            {
                AutoScrollPosition = Point.Empty;
                Invalidate();
                return;
            }

            AutoScrollPosition = recenterViewport
                ? GetCenteredScrollOffset(contentSize)
                : ClampScrollOffset(scrollOffset, contentSize);
            Invalidate();
        }


        private void ApplyLayout(PointF sourcePoint, Point focusPoint, bool recenterViewport)
        {
            Size contentSize = BackgroundImage != null ? GetContentSize(m_zoomFactor) : ClientSize;
            AutoScrollMinSize = contentSize;

            if (BackgroundImage == null)
            {
                AutoScrollPosition = Point.Empty;
                Invalidate();
                return;
            }

            Point scrollOffset;
            if (recenterViewport)
            {
                scrollOffset = GetCenteredScrollOffset(contentSize);
            }
            else
            {
                Size scaledImageSize = GetScaledImageSize(m_zoomFactor);
                Point imageOrigin = GetImageOrigin(contentSize, scaledImageSize);
                scrollOffset = new Point(
                    (int)Math.Round(imageOrigin.X + sourcePoint.X * m_zoomFactor - focusPoint.X),
                    (int)Math.Round(imageOrigin.Y + sourcePoint.Y * m_zoomFactor - focusPoint.Y)
                );
                scrollOffset = ClampScrollOffset(scrollOffset, contentSize);
            }

            AutoScrollPosition = scrollOffset;
            Invalidate();
        }


        protected override void OnBackgroundImageChanged(EventArgs e)
        {
            base.OnBackgroundImageChanged(e);
            UpdateLayout(true);
        }


        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            UpdateLayout(false);
        }


        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);

            if (EnableInteractiveZoom && CanFocus && !Focused)
            {
                Focus();
            }
        }


        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (EnableInteractiveZoom && CanFocus && !Focused)
            {
                Focus();
            }

            if (EnableInteractiveZoom && BackgroundImage != null && e.Button == MouseButtons.Middle)
            {
                ResetZoom();
                return;
            }

            base.OnMouseDown(e);
        }


        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if (EnableInteractiveZoom && BackgroundImage != null && (ModifierKeys & Keys.Control) == Keys.Control)
            {
                int wheelSteps = e.Delta / SystemInformation.MouseWheelScrollDelta;
                if (wheelSteps != 0)
                {
                    int currentZoomStep = (int)Math.Round(m_zoomFactor / ZoomStep, MidpointRounding.AwayFromZero);
                    float newZoom = (currentZoomStep + wheelSteps) * ZoomStep;
                    SetZoomFactor(newZoom, e.Location, false);
                }
                return;
            }

            base.OnMouseWheel(e);
        }


        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(SystemColors.Control);

            Size contentSize = BackgroundImage != null ? GetContentSize(m_zoomFactor) : ClientSize;

            // draw background
            if (Background == BackgroundStyle.Chessboard)
            {
                int numXCells = contentSize.Width / m_cellSize.Width + 1;
                int numYCells = contentSize.Height / m_cellSize.Height + 1;

                e.Graphics.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);
                for (int x = 0; x < numXCells; ++x)
                {
                    for (int y = 0; y < numYCells; ++y)
                    {
                        Brush brush;
                        if ((x + y) % 2 > 0)
                            brush = Brushes.White;
                        else brush = Brushes.LightGray;

                        e.Graphics.FillRectangle(brush,
                            new Rectangle(
                                x * m_cellSize.Width,
                                y * m_cellSize.Height,
                                m_cellSize.Width,
                                m_cellSize.Height)
                        );
                    }
                }
                e.Graphics.TranslateTransform(-AutoScrollPosition.X, -AutoScrollPosition.Y);
            }
            else
            {
                using (Brush brush = new SolidBrush(BackColor))
                {
                    e.Graphics.FillRectangle(brush, ClientRectangle);
                }
            }

            // draw image
            if (BackgroundImage == null)
            {
                return;
            }

            Rectangle imageBounds = GetImageBounds(m_zoomFactor);
            InterpolationMode previousInterpolation = e.Graphics.InterpolationMode;
            e.Graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
            e.Graphics.DrawImage(BackgroundImage, imageBounds);
            e.Graphics.InterpolationMode = previousInterpolation;

            if (HighlightArea != null)
            {
                Rectangle scaledHighlight = ScaleRectangle(HighlightArea.Value, m_zoomFactor);
                scaledHighlight.Offset(imageBounds.Location);
                int contentLeft = AutoScrollPosition.X;
                int contentTop = AutoScrollPosition.Y;
                int contentRight = contentLeft + contentSize.Width;
                int contentBottom = contentTop + contentSize.Height;

                using (Pen p = new Pen(Color.Violet, 2.0f))
                {
                    e.Graphics.DrawLine(p, scaledHighlight.Left, contentTop, scaledHighlight.Left, contentBottom);
                    e.Graphics.DrawLine(p, scaledHighlight.Right, contentTop, scaledHighlight.Right, contentBottom);

                    e.Graphics.DrawLine(p, contentLeft, scaledHighlight.Top, contentRight, scaledHighlight.Top);
                    e.Graphics.DrawLine(p, contentLeft, scaledHighlight.Bottom, contentRight, scaledHighlight.Bottom);
                }
            }
        }
    }
}

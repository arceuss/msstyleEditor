using libmsstyle;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msstyleEditor
{
    class PartRenderer
    {
        private VisualStyle m_style;
        private StylePart m_part;
        private StyleState m_selectedState;
        private StyleState m_commonState;

        public PartRenderer(VisualStyle style, StylePart part, StyleState selectedState = null)
        {
            m_style = style;
            m_part = part;
            m_selectedState = selectedState;
            m_part.States.TryGetValue(0, out m_commonState);
        }

        public Bitmap RenderPreview()
        {
            if (GetPrimaryState() == null)
            {
                return null;
            }

            Bitmap surface = new Bitmap(150, 50, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(surface))
            {
                DrawBackground(g);
            }
            
            return surface;
        }

        private StyleState GetPrimaryState()
        {
            if (m_selectedState != null)
            {
                return m_selectedState;
            }

            return m_commonState;
        }

        private IEnumerable<StyleState> GetPropertySearchStates()
        {
            if (m_selectedState != null)
            {
                yield return m_selectedState;
            }

            if (m_commonState != null && !ReferenceEquals(m_commonState, m_selectedState))
            {
                yield return m_commonState;
            }
        }

        private bool TryGetPropertyValue<T>(IDENTIFIER ident, ref T value)
        {
            foreach (var state in GetPropertySearchStates())
            {
                if (state.TryGetPropertyValue(ident, ref value))
                {
                    return true;
                }
            }

            return false;
        }

        private StyleProperty FindProperty(params IDENTIFIER[] propertyIds)
        {
            foreach (var state in GetPropertySearchStates())
            {
                foreach (var propertyId in propertyIds)
                {
                    var prop = state.Properties.Find((p) => p.Header.nameID == (int)propertyId);
                    if (prop != null)
                    {
                        return prop;
                    }
                }
            }

            return null;
        }

        private bool TryGetAtlasRect(StyleState state, out Rectangle atlasRect)
        {
            atlasRect = Rectangle.Empty;
            if (state == null)
            {
                return false;
            }

            var rectProp = state.Properties.Find((p) => p.Header.nameID == (int)IDENTIFIER.ATLASRECT);
            if (rectProp == null)
            {
                return false;
            }

            var mt = rectProp.GetValueAs<Margins>();
            atlasRect = Rectangle.FromLTRB(mt.Left, mt.Top, mt.Right, mt.Bottom);
            return atlasRect.Width > 0 && atlasRect.Height > 0;
        }

        private Rectangle SliceImageRegion(Rectangle imageRegion, int imageLayout, int imageCount, int imageIndex)
        {
            int left = imageRegion.Left;
            int top = imageRegion.Top;
            int right = imageRegion.Right;
            int bottom = imageRegion.Bottom;

            if (imageLayout == 1)
            {
                left += imageRegion.Width * imageIndex / imageCount;
                right = imageRegion.Left + imageRegion.Width * (imageIndex + 1) / imageCount;
            }
            else
            {
                top += imageRegion.Height * imageIndex / imageCount;
                bottom = imageRegion.Top + imageRegion.Height * (imageIndex + 1) / imageCount;
            }

            return Rectangle.FromLTRB(left, top, right, bottom);
        }

        private bool TryGetPreviewImageIndex(int imageCount, out int imageIndex)
        {
            imageIndex = 0;
            if (imageCount <= 1)
            {
                return false;
            }

            if (m_selectedState != null && m_selectedState.StateId > 0)
            {
                imageIndex = Math.Min(m_selectedState.StateId - 1, imageCount - 1);
                return true;
            }

            imageIndex = 0;
            return true;
        }

        private Rectangle ResolveImageRegion(Image fullImage)
        {
            Rectangle imageRegion;
            if (TryGetAtlasRect(m_selectedState, out imageRegion))
            {
                imageRegion.Intersect(new Rectangle(0, 0, fullImage.Width, fullImage.Height));
                return imageRegion;
            }

            if (TryGetAtlasRect(m_commonState, out imageRegion))
            {
                int imageCount = 1;
                TryGetPropertyValue(IDENTIFIER.IMAGECOUNT, ref imageCount);

                int imageIndex;
                if (TryGetPreviewImageIndex(imageCount, out imageIndex))
                {
                    int imageLayout = 0;
                    TryGetPropertyValue(IDENTIFIER.IMAGELAYOUT, ref imageLayout);
                    imageRegion = SliceImageRegion(imageRegion, imageLayout, imageCount, imageIndex);
                }

                imageRegion.Intersect(new Rectangle(0, 0, fullImage.Width, fullImage.Height));
                return imageRegion;
            }

            imageRegion = new Rectangle(0, 0, fullImage.Width, fullImage.Height);

            int fallbackImageCount = 1;
            TryGetPropertyValue(IDENTIFIER.IMAGECOUNT, ref fallbackImageCount);
            int fallbackImageIndex;
            if (TryGetPreviewImageIndex(fallbackImageCount, out fallbackImageIndex))
            {
                int fallbackImageLayout = 0;
                TryGetPropertyValue(IDENTIFIER.IMAGELAYOUT, ref fallbackImageLayout);
                imageRegion = SliceImageRegion(imageRegion, fallbackImageLayout, fallbackImageCount, fallbackImageIndex);
            }

            imageRegion.Intersect(new Rectangle(0, 0, fullImage.Width, fullImage.Height));
            return imageRegion;
        }

        private void DrawBackground(Graphics g)
        {
            int bgType = 2;
            TryGetPropertyValue(IDENTIFIER.BGTYPE, ref bgType);

            switch(bgType)
            {
                case 0: // IMAGEFILL
                    DrawBackgroundImageFill(g); break;
                case 1: // BORDERFILL
                    DrawBackgroundSolidFill(g); break;
                default: break;
            }
        }

        private void DrawBackgroundImageFill(Graphics g)
        {
            var imageFileProp = FindProperty(IDENTIFIER.IMAGEFILE, IDENTIFIER.IMAGEFILE1);
            if (imageFileProp == null)
            {
                return;
            }

            var resource = m_style.GetResourceFromProperty(imageFileProp);
            if (resource?.Data == null)
            {
                return;
            }

            StyleResourceType resourceType = resource.Type;
            string file = m_style.GetQueuedResourceUpdate(imageFileProp.Header.shortFlag, resourceType);

            using (Image fullImage = !String.IsNullOrEmpty(file)
                ? Image.FromFile(file)
                : Image.FromStream(new MemoryStream(resource.Data)))
            {
                Rectangle imagePartToDraw = ResolveImageRegion(fullImage);
                if (imagePartToDraw.Width <= 0 || imagePartToDraw.Height <= 0)
                {
                    return;
                }

                int sizingType = 0; // TRUESIZE
                TryGetPropertyValue(IDENTIFIER.SIZINGTYPE, ref sizingType);

                Margins sizingMargins = default(Margins);
                bool haveSizingMargins = TryGetPropertyValue(IDENTIFIER.SIZINGMARGINS, ref sizingMargins);
                bool sizingMarginsZero = new Margins(0, 0, 0, 0).Equals(sizingMargins);

                switch (sizingType)
                {
                    case 0: // TRUESIZE
                        {
                            Rectangle dst = new Rectangle(
                                (int)(g.VisibleClipBounds.Width / 2) - (imagePartToDraw.Width / 2),
                                (int)(g.VisibleClipBounds.Height / 2) - (imagePartToDraw.Height / 2),
                                imagePartToDraw.Width, imagePartToDraw.Height);

                            g.DrawImage(fullImage, dst, imagePartToDraw, GraphicsUnit.Pixel);
                        }
                        break;
                    case 1: // STRETCH
                    case 2: // TILE
                        {
                            if (haveSizingMargins && !sizingMarginsZero)
                            {
                                Rectangle absMargins = new Rectangle(
                                    sizingMargins.Left,
                                    sizingMargins.Top,
                                    Math.Max(imagePartToDraw.Width - sizingMargins.Left - sizingMargins.Right, 1),
                                    Math.Max(imagePartToDraw.Height - sizingMargins.Top - sizingMargins.Bottom, 1)
                                );

                                DrawImage9SliceScaled(g, fullImage, imagePartToDraw, Rectangle.Round(g.VisibleClipBounds), absMargins);
                            }
                            else
                            {
                                using (ImageAttributes attr = new ImageAttributes())
                                {
                                    attr.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                                    g.DrawImage(fullImage, Rectangle.Round(g.VisibleClipBounds),
                                        imagePartToDraw.X,
                                        imagePartToDraw.Y,
                                        imagePartToDraw.Width,
                                        imagePartToDraw.Height,
                                        GraphicsUnit.Pixel, attr);
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private void DrawBackgroundSolidFill(Graphics g)
        {
            Color bgFill = Color.White;
            if (TryGetPropertyValue(IDENTIFIER.FILLCOLOR, ref bgFill))
            {
                using (var brush = new SolidBrush(bgFill))
                {
                    g.FillRectangle(brush, g.ClipBounds);
                }
            }
        }

        private void DrawImage9SliceScaled(Graphics g, Image image, Rectangle src, Rectangle dst, Rectangle sm)
        {
            /* SRC */
            /* Assumes that sm is adjusted to this image */
            Rectangle topLeft   = new Rectangle(src.X           , src.Y, sm.Left                , sm.Top);
            Rectangle topMiddle = new Rectangle(src.X + sm.Left , src.Y, sm.Width               , sm.Top);
            Rectangle topRight  = new Rectangle(src.X + sm.Right, src.Y, src.Width - sm.Right   , sm.Top);

            Rectangle midLeft   = new Rectangle(src.X           , src.Y + sm.Top, sm.Left                   , sm.Height);
            Rectangle midMiddle = new Rectangle(src.X + sm.Left , src.Y + sm.Top, sm.Width                  , sm.Height);
            Rectangle midRight  = new Rectangle(src.X + sm.Right, src.Y + sm.Top, src.Width - sm.Right      , sm.Height);

            Rectangle botLeft   = new Rectangle(src.X           , src.Y + sm.Bottom, sm.Left                    , src.Height - sm.Bottom);
            Rectangle botMiddle = new Rectangle(src.X + sm.Left , src.Y + sm.Bottom, sm.Width                   , src.Height - sm.Bottom);
            Rectangle botRight  = new Rectangle(src.X + sm.Right, src.Y + sm.Bottom, src.Width - sm.Right       , src.Height - sm.Bottom);

            /* DST */
            int varWidth = dst.Width - sm.Left - (src.Width - sm.Right);
            int varHeight = dst.Height - sm.Top - (src.Height - sm.Bottom);

            Rectangle topLeftDst    = new Rectangle(dst.X                  , dst.Y, topLeft.Width , topLeft.Height);      // keep W & H
            Rectangle topMiddleDst  = new Rectangle(dst.X + sm.Left        , dst.Y, varWidth      , topMiddle.Height);    // keep H
            Rectangle topRightDst   = new Rectangle(dst.X + sm.Left + varWidth, dst.Y, topRight.Width, topRight.Height);     // keep W & H

            Rectangle midLeftDst    = new Rectangle(dst.X                  , dst.Y + sm.Top, midLeft.Width, varHeight); // keep W
            Rectangle midMiddleDst  = new Rectangle(dst.X + sm.Left        , dst.Y + sm.Top, varWidth     , varHeight); // fully scale
            Rectangle midRightDst   = new Rectangle(dst.X + sm.Left + varWidth, dst.Y + sm.Top, midRight.Width, varHeight); // keep W

            Rectangle botLeftDst    = new Rectangle(dst.X                  , dst.Y + sm.Top + varHeight, botLeft.Width    , botLeft.Height);      // keep W & H
            Rectangle botMiddleDst  = new Rectangle(dst.X + sm.Left        , dst.Y + sm.Top + varHeight, varWidth         , botMiddle.Height);    // keep H
            Rectangle botRightDst   = new Rectangle(dst.X + sm.Left + varWidth, dst.Y + sm.Top + varHeight, botRight.Width   , botRight.Height);     // keep W & H

            g.DrawImage(image, topLeftDst, topLeft, GraphicsUnit.Pixel);
            g.DrawImage(image, topMiddleDst, topMiddle, GraphicsUnit.Pixel);
            g.DrawImage(image, topRightDst, topRight, GraphicsUnit.Pixel);

            g.DrawImage(image, midLeftDst, midLeft, GraphicsUnit.Pixel);
            g.DrawImage(image, midMiddleDst, midMiddle, GraphicsUnit.Pixel);
            g.DrawImage(image, midRightDst, midRight, GraphicsUnit.Pixel);

            g.DrawImage(image, botLeftDst, botLeft, GraphicsUnit.Pixel);
            g.DrawImage(image, botMiddleDst, botMiddle, GraphicsUnit.Pixel);
            g.DrawImage(image, botRightDst, botRight, GraphicsUnit.Pixel);
        }
    }
}

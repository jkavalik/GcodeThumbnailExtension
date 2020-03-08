using System;
using System.Drawing;
using SharpShell.SharpThumbnailHandler;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text;
using System.Runtime.InteropServices;
using SharpShell.Attributes;

namespace GcodeThumbnailExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".gcode", ".gco")]
    public class CgodeThumbnailHandler : SharpThumbnailHandler
    {
        protected override Bitmap GetThumbnailImage(uint width)
        {
            //  Create a stream reader for the selected item stream
            try
            {
                using (var reader = new StreamReader(SelectedItemStream))
                {
                    //  Now return a preview from the gcode or error when none present
                    return GetThumbnailForGcode(reader, width);
                }
            }
            catch (Exception exception)
            {
                //  Log the exception and return null for failure (no thumbnail to show)
                LogError("An exception occurred opening the file.", exception);
                return null;
            }
        }

        /// <summary>
        /// Create the resized bitmap representation that can be used by the Shell
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        private Bitmap GetThumbnailForGcode(StreamReader reader, uint width)
        {
            //  Create the bitmap dimensions
            var thumbnailSize = new Size((int)width, (int)width);

            //  Create the bitmap
            var bitmap = new Bitmap(thumbnailSize.Width, thumbnailSize.Height,
                                    PixelFormat.Format32bppArgb);

            // Get pre-generated thumbnail from gcode file or error if none found
            Image gcodeThumbnail = ReadThumbnailFromGcode(reader);

            //  Create a graphics object to render to the bitmap
            using (var graphics = Graphics.FromImage(bitmap))
            {
                //  Set the rendering up for anti-aliasing
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.DrawImage(gcodeThumbnail, 0, 0, thumbnailSize.Width, thumbnailSize.Height);
            }

            //  Return the bitmap
            return bitmap;
        }

        /// <summary>
        /// Parse the gcode recursively to find the last thumbnail and return the decoded image
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private Image ReadThumbnailFromGcode(StreamReader reader)
        {
            int thumbnailSize = 0; // currently unused, intended to enable picking the best thumbnail image later
            int counter = 0, maxCounter = 1000;
            String line = "";

            // locate first thumbnail block in the header
            while ((counter < maxCounter) && !reader.EndOfStream) {
                counter++;
                line = reader.ReadLine();
                if (line.Contains("thumbnail begin"))
                {
                    break;
                }
            }

            // if no thumbnail was found in first maxCounter lines then throw error
            if ((counter == maxCounter) || reader.EndOfStream)
            {
                throw new Exception();
            }

            // parse the thumbnail size (it seems somehow imprecise, the number being slightly smaller than actual base64 encoded file size
            String[] thumbnailMeta = line.Split(' ');
            thumbnailSize = Int32.Parse(thumbnailMeta[thumbnailMeta.Length - 1]);

            // load current thumbnail data
            StringBuilder thumbnailData = ReadBase64Data(reader);
            byte[] imageBytes = Convert.FromBase64String(thumbnailData.ToString());

            // try to find additional thumbnails in the file (PrusaSlicer seems to order them by ascending resolution)
            try
            {
                Image otherImg = ReadThumbnailFromGcode(reader);
                return otherImg;
            }
            catch (Exception)
            {
                // do nothing, we got the current thumbnail anyway
            }

            // if no later thumbnail was found or its decoding ended with an error
            // we load the current one to image instance
            using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
            {
                Image image = Image.FromStream(ms, true);
                return image;
            }
        }

        /// <summary>
        /// Read base64 data (present as comments) from the current position of the given stream until first "thumbnail end"
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private StringBuilder ReadBase64Data(StreamReader reader)
        {
            StringBuilder res = new StringBuilder(20000);
            String line = reader.ReadLine();
            // read all lines for current thumbnail
            while ((!line.Contains("thumbnail end")) && !reader.EndOfStream)
            {
                // trim whitespaces and semicolon - the thumbnail is embedded as a comment
                res.Append(line.Trim("; \n".ToCharArray()));
                line = reader.ReadLine();
            }

            // no thumbnail end found, file is damaged or invalid
            if (reader.EndOfStream)
            {
                throw new Exception();
            }
            return res;
        }
    }
}

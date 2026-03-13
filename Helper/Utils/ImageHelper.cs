using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace BatDongSan.Utils
{
    public static class ImageHelper
    {
        /// <summary>
        /// Checking whether the image needs to be resized
        /// </summary>
        /// <param name="uploadImage"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <returns>true or false depending on the size</returns>
        private static bool IsResizeNeeded(Image uploadImage, int height, int width)
        {
            var originalWidth = uploadImage.Width;
            var originalHeight = uploadImage.Height;
            return (originalHeight != height) || (originalWidth != width);
        }

        /// <summary>
        /// Resize the image depending on the size given
        /// </summary>
        /// <param name="uploadImage"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        public static Image ResizeBySize(Image uploadImage, int height, int width)
        {
            var originalWidth = uploadImage.Width;
            var originalHeight = uploadImage.Height;

            var modifiedHeight = height;
            var modifiedWidth = width;

            var bitmap = new Bitmap(modifiedWidth, modifiedHeight,
                                     PixelFormat.Format32bppPArgb);
            bitmap.SetResolution(uploadImage.HorizontalResolution, uploadImage.VerticalResolution);

            var graphics = Graphics.FromImage(bitmap);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

            graphics.DrawImage(uploadImage,
                              new Rectangle(0, 0, modifiedWidth, modifiedHeight),
                              new Rectangle(0, 0, originalWidth, originalHeight),
                              GraphicsUnit.Pixel);

            graphics.Dispose();

            return bitmap;
        }

        /// <summary>
        /// Since we need the ImageFormat inorder to convert Image -> MemoryStream
        /// </summary>
        /// <param name="imageType"></param>
        /// <returns>ImageFormat according to the image type(Ex: jpg, jpeg etc.)</returns>
        private static ImageFormat GetImageFormat(string imageType)
        {
            ImageFormat imageFormat;
            switch (imageType)
            {
                case "image/jpg":
                    imageFormat = ImageFormat.Jpeg;
                    break;
                case "image/jpeg":
                    imageFormat = ImageFormat.Jpeg;
                    break;
                case "image/pjpeg":
                    imageFormat = ImageFormat.Jpeg;
                    break;
                case "image/gif":
                    imageFormat = ImageFormat.Gif;
                    break;
                case "image/png":
                    imageFormat = ImageFormat.Png;
                    break;
                case "image/x-png":
                    imageFormat = ImageFormat.Png;
                    break;
                default:
                    throw new Exception("Unsupported image type !");
            }

            return imageFormat;
        }

        /// <summary>
        /// Convert byte array to image
        /// </summary>
        /// <param name="byteArrayIn"></param>
        /// <returns></returns>
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }

        /// <summary>
        /// Convert files to byte array(Special function added for images which needs to be resized)
        /// </summary>
        /// <param name="fileUpLoad"></param>
        /// <returns>byte array of the image file</returns>
        public static byte[] ToByteArrayResized(this HttpPostedFileBase fileUpLoad)
        {
            byte[] byteArray;

            if (IsImage(fileUpLoad))
            {
                //getting the height and width of the image to be uploaded
                var theHeight = Int32.Parse(WebConfigurationManager.AppSettings["UploadImageHeight"]);
                var theWidth = Int32.Parse(WebConfigurationManager.AppSettings["UploadImageWidth"]);

                //converting to a bitmap
                var uploadImage = new Bitmap(fileUpLoad.InputStream);

                //checking whether resize is needed
                if (IsResizeNeeded(uploadImage, theHeight, theWidth))
                {
                    //resizing the image
                    var resizedImage = ResizeBySize(uploadImage, theHeight, theWidth);
                    //getting the image format(not the type)
                    var resizedImageFormat = GetImageFormat(fileUpLoad.ContentType);

                    using (resizedImage)
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            //loading it to the memory stream
                            if (resizedImageFormat != null)
                            {
                                resizedImage.Save(memoryStream, resizedImageFormat);
                            }
                            memoryStream.Position = 0;
                            byteArray = memoryStream.ToArray();
                        }
                    }
                }
                else
                {
                    byteArray = ByteArrayConvertion(fileUpLoad);
                }
            }
            // for non-image files
            else
            {
                byteArray = ByteArrayConvertion(fileUpLoad);
            }

            return byteArray;
        }

        /// <summary>
        /// Convert files to byte array
        /// </summary>
        /// <param name="fileUpLoad"></param>
        /// <returns>byte array of the image file</returns>
        public static byte[] ToByteArray(this HttpPostedFileBase fileUpLoad)
        {
            byte[] byteArray;

            if (IsImage(fileUpLoad))
            {
                //converting to a bitmap
                var uploadImage = new Bitmap(fileUpLoad.InputStream);
                byteArray = ByteArrayConvertion(fileUpLoad);
            }
            // for non-image files
            else
            {
                byteArray = null;
            }
            return byteArray;
        }


        /// <summary>
        /// Default byte array creation
        /// </summary>
        /// <param name="fileUpload"></param>
        /// <returns>byte array</returns>
        private static byte[] ByteArrayConvertion(HttpPostedFileBase fileUpload)
        {
            byte[] byteArray;

            using (fileUpload.InputStream)
            {
                using (var memoryStream = new MemoryStream())
                {
                    fileUpload.InputStream.CopyTo(memoryStream);
                    byteArray = memoryStream.ToArray();
                }
            }
            return byteArray;
        }

        /// <summary>
        /// Checking the file is an image....
        /// </summary>
        /// <param name="fileUpload"></param>
        /// <returns>True if the file is an image</returns>
        /// http://yassershaikh.com/how-to-check-if-an-uploaded-file-is-an-image-or-not-in-asp-net-mvc-3/
        private static bool IsImage(HttpPostedFileBase fileUpload)
        {
            if (fileUpload.ContentType.Contains("image"))
            {
                return true;
            }

            var formats = new string[] { ".jpg", ".png", ".gif", ".jpeg" }; // add more if u like...

            // linq from Henrik Stenbæk
            return formats.Any(item => fileUpload.FileName.EndsWith(item, StringComparison.OrdinalIgnoreCase));
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

        public static void CheckServerMap(string path)
        {
            if (!Directory.Exists(path))
            {
                //If No any such directory then creates the new one
                Directory.CreateDirectory(path);
            }
        }
    }
}
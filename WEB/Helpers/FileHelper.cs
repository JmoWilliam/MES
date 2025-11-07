using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Helpers
{
    public class FileHelper
    {
        public static string GetMime(string extension)
        {
            return MimeMapping.GetMimeMapping(extension);
        }

        public static List<FileModel> FileSave(HttpFileCollectionBase uploadFiles)
        {
            List<FileModel> uploadFile = new List<FileModel>();

            for (int i = 0; i < uploadFiles.Count; i++)
            {
                BinaryReader binaryReader = new BinaryReader(uploadFiles[i].InputStream);

                uploadFile.Add(
                    new FileModel
                    {
                        FileName = Path.GetFileNameWithoutExtension(uploadFiles[i].FileName),
                        FileContent = binaryReader.ReadBytes(uploadFiles[i].ContentLength),
                        FileExtension = Path.GetExtension(uploadFiles[i].FileName),
                        FileSize = uploadFiles[i].ContentLength
                    });
            }

            return uploadFile;
        }

        public static void FileType(ref string filetype, ref bool filePreview, string extension)
        {
            switch (extension)
            {                
                case ".gif":
                case ".jpg":
                case ".jpeg":
                case ".png":
                    filetype = "image";
                    filePreview = true;
                    break;
                case ".htm":
                case ".html":
                    filetype = "html";
                    filePreview = true;
                    break;
                case ".cs":
                case ".css":
                case ".ini":
                case ".js":
                case ".json":
                case ".md":
                case ".nfo":
                case ".php":                
                case ".txt":
                    filetype = "text";
                    filePreview = true;
                    break;
                case ".doc":
                case ".docx":
                case ".potx":
                case ".pps":
                case ".ppt":
                case ".pptx":
                case ".xls":
                case ".xlsx":
                    filetype = "other";
                    filePreview = true;
                    break;
                case ".sql":
                    filetype = "text";
                    filePreview = true;
                    break;
                case ".ai":
                case ".dxf":
                case ".eps":
                case ".ods":
                case ".odt":
                case ".pages":
                case ".rtf":
                case ".tif":
                case ".tiff":
                case ".ttf":
                case ".wmf":
                    filetype = "gdocs";
                    filePreview = false;
                    break;
                case ".3gp":
                case ".mov":
                case ".mp4":
                case ".mpg":
                case ".ogg":
                case ".webm":
                    filetype = "video";
                    filePreview = true;
                    break;
                case ".og":
                case ".mp3":
                case ".wav":
                    filetype = "audio";
                    filePreview = true;
                    break;                
                case ".swf":
                    filetype = "flash";
                    filePreview = true;
                    break;
                case ".pdf":
                    filetype = "pdf";
                    filePreview = true;
                    break;
                default:
                    filetype = "other";
                    filePreview = true;
                    break;
            }
        }
    }
}

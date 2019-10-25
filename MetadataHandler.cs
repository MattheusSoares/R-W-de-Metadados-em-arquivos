using System;
using System.IO;
using DSOFile;
using Microsoft.WindowsAPICodePack.Shell;
using TagLib;

namespace DELE
{
    class MetadataHandler
    {
        public string FilePath { get; private set; }
        public string FileExtension { get; private set; }
        public string Metadata { get; private set; }

        public MetadataHandler(string filePath)
        {
            FilePath = filePath;
            FileExtension = Path.GetExtension(filePath).Substring(1).ToLower();
        }

        public void WriteFileMetadata(string metadata) {

            if (VerifyFileExtension() == "media")
            {
                MediaMetadataHandlerWriter(metadata);
            }
            else if (VerifyFileExtension() == "doc")
            {
                DocMetadataHandlerWriter(metadata);
            }
        }

        public string ReadFileMetadata()
        {
            string FileMetadata;

            if (VerifyFileExtension() == "media")
            {
                FileMetadata = MediaMetadataHandlerReader();
            }
            else if (VerifyFileExtension() == "doc")
            {
                FileMetadata = DocMetadataHandlerReader();
            }
            else
            {
                FileMetadata = "null";
            }

            return FileMetadata;
        }


        private string VerifyFileExtension()
        {
            string fileType = "";

            if (FileExtension == "jpg" || FileExtension == "jpeg" || FileExtension == "tiff" || FileExtension == "aac" ||
                FileExtension == "mp3" || FileExtension == "flac" || FileExtension == "ogg" || FileExtension == "wav" || 
                FileExtension == "wma" || FileExtension == "avi" || FileExtension == "mov" || FileExtension == "mp4" ||
                FileExtension == "wmv")
            {
                fileType = "media";
            }
            else if (FileExtension == "pdf" || FileExtension == "doc" || FileExtension == "xls" || FileExtension == "ppt" ||
                     FileExtension == "docx" || FileExtension == "xlsx" || FileExtension == "pptx" || FileExtension == "json" ||
                     FileExtension == "odt" || FileExtension == "ods" || FileExtension == "odp" || FileExtension == "xml" ||
                     FileExtension == "7z" || FileExtension == "zip" || FileExtension == "rar" || FileExtension == "iso" ||
                     FileExtension == "rtf" || FileExtension == "txt" || FileExtension == "cdr" || FileExtension == "eps" ||
                     FileExtension == "psd" || FileExtension == "tga" || FileExtension == "css" || FileExtension == "js" ||
                     FileExtension == "cs" || FileExtension == "cshtml" || FileExtension == "csproj" || FileExtension == "dwg" ||
                     FileExtension == "exe" || FileExtension == "frm" || FileExtension == "frx" || FileExtension == "mdb" ||
                     FileExtension == "swf" || FileExtension == "vbp" || FileExtension == "vbw" || FileExtension == "vob" || 
                     FileExtension == "img" || FileExtension == "fla" || FileExtension == "htm" || FileExtension == "bas" ||
                     FileExtension == "cls" || FileExtension == "vb" || FileExtension == "vbproj")
            {
                fileType = "doc";
            }
            else
            {
                fileType = "unsupported";
            }

            return fileType;
        }

        private string MediaMetadataHandlerReader()
        {
            TagLib.File file = TagLib.File.Create(FilePath);

            Metadata = file.Tag.Copyright.ToString();

            return Metadata;
        }

        private void MediaMetadataHandlerWriter(string metadata)
        {
            TagLib.File file = TagLib.File.Create(FilePath);

            file.Tag.Copyright = metadata;

            file.Save();

            //return file;
        }

        private string DocMetadataHandlerReader()
        {
            OleDocumentProperties File = new OleDocumentProperties();
            File.Open(FilePath, false, dsoFileOpenOptions.dsoOptionDefault);

            CustomProperties customProperties = File.CustomProperties;

            object index = "Copyright";

            Metadata = "";

            foreach (CustomProperty property in File.CustomProperties)
            {
                if (property.Name == index.ToString())
                {
                    Metadata = customProperties[index].get_Value();
                    break;
                }
            }

            return Metadata;
        }

        private void DocMetadataHandlerWriter(string metadata)
        {
            OleDocumentProperties File = new OleDocumentProperties();
            File.Open(FilePath, false, dsoFileOpenOptions.dsoOptionDefault);

            string propertyName = "Copyright";
            object propertyValue = metadata;

            bool hasPropertyFlag = false;

            foreach (CustomProperty property in File.CustomProperties)
            {
                if (property.Name == propertyName)
                {
                    hasPropertyFlag = true;
                    property.set_Value(propertyValue);
                    break;
                }
            }

            if (!hasPropertyFlag)
            {
                File.CustomProperties.Add(propertyName, ref propertyValue);
            }

            File.Save();
            File.Close(true);
            //return File;
        }
    }
}

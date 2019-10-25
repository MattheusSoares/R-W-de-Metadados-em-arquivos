using System;
using System.IO;
using DSOFile;
using Microsoft.WindowsAPICodePack.Shell;
using TagLib;

namespace DELE
{
    class Program
    {
        static void Main(string[] args)
        {

            string path = @"C:\z - samples\samples\sample ok.psd";
            MetadataHandler metadataHandler = new MetadataHandler(path);
            metadataHandler.WriteFileMetadata("TESTE classe");
            Console.WriteLine(metadataHandler.ReadFileMetadata());
            Console.ReadKey();




            /*
             * Mídia (Vídeo, Áudio e Imagem)
             */
            //string path = @"C:\z - samples\samples\sample.mpg";

            //TagLib.File file = TagLib.File.Create(path);

            ////TagLib.Image.File image = file as TagLib.Image.File;

            ////image.ImageTag.Copyright = "Teste para Copyright de Metadata";
            //file.Tag.Copyright = "Teste para Copyright de Metadata";

            //file.Save();
            ////image.Save();

            //Console.WriteLine(file.Tag.Copyright.ToString());

            //Console.ReadKey();





            /*
             * Outros Docs
             */

            //string path = @"C:\z - samples\samples\sample.sln";

            //OleDocumentProperties myFile = new OleDocumentProperties();
            //myFile.Open(path, false, dsoFileOpenOptions.dsoOptionDefault);

            //object yourValue = "Teste para Copyright de Metadata";

            //string propertyName = "Copyright";
            //foreach (CustomProperty property in myFile.CustomProperties)
            //{
            //    if (property.Name == propertyName)
            //    {
            //        property.set_Value(yourValue);
            //    }
            //}

            //myFile.CustomProperties.Add(propertyName, ref yourValue);

            //CustomProperties customProperties = myFile.CustomProperties;

            //object index = "Copyright";

            //Console.WriteLine(customProperties[index].get_Value());
            //Console.ReadKey();

            //myFile.Save();
            //myFile.Close(true);
        }

        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Ionic.Zip;
using OpenPop.Mime;

namespace LibraryOutlook.SubscribeOutlook
{
   public class ZipAttachments
    {
        /// <summary>
        /// Архивирование модели POP3 Файлов
        /// </summary>
        /// <param name="collectionMessageAttach"></param>
        public byte[] StartZipArchive(List<MessagePart> collectionMessageAttach,string nameZip)
        {
            if (collectionMessageAttach.Count > 0)
            {
                using (var zip = File.Open(nameZip, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    using (var zipToWrite = new ZipOutputStream(zip))
                    {
                        var i = 0;
                        foreach (var messageAttach in collectionMessageAttach)
                        {
                            var fileStream = messageAttach.Body;
                            var nameFile = $"{i}_{messageAttach.FileName}";
                            using (Stream newFileStream =new MemoryStream(fileStream))
                            {
                                var byteBuffer = new byte[newFileStream.Length - 1];
                                newFileStream.Read(byteBuffer, 0, byteBuffer.Length);
                                zipToWrite.EnableZip64 = Zip64Option.Never;
                                zipToWrite.AlternateEncodingUsage = ZipOption.Always;
                                zipToWrite.AlternateEncoding = Encoding.GetEncoding(866);
                                zipToWrite.PutNextEntry(nameFile);
                                zipToWrite.Write(byteBuffer, 0, byteBuffer.Length);
                            }
                            i++;
                        }
                    }
                    zip.Close();
                }
                return File.ReadAllBytes(nameZip);
            }
            return null;
        }
    }
}

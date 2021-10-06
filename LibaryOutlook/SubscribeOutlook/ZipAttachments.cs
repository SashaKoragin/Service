using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Ionic.Zip;
using MimeKit;
using MessagePart = MimeKit.MessagePart;

namespace LibraryOutlook.SubscribeOutlook
{
   public class ZipAttachments
    {
        /// <summary>
        /// Архивирование модели POP3 Файлов
        /// </summary>
        /// <param name="collectionMessageAttach">Коллекция пришедшей почты по POP3</param>
        /// <param name="fullPathZip">Полный путь к наименованию Архива файла</param>
        public byte[] StartZipArchive(List<MimeEntity> collectionMessageAttach,string fullPathZip)
        {
            if (collectionMessageAttach.Any())
            {
                using (var zip = File.Open(fullPathZip, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    using (var zipToWrite = new ZipOutputStream(zip))
                    {
                        var i = 0;
                        foreach (var messageAttach in collectionMessageAttach)
                        {
                          var nameFile = $"{i}_{messageAttach.ContentDisposition?.FileName}";
                          using (var memory = new MemoryStream())
                          {
                              if (messageAttach is MimePart)
                                  ((MimePart)messageAttach).Content.DecodeTo(memory);
                              else
                                  ((MessagePart)messageAttach).Message.WriteTo(memory);
                              var bytes = memory.ToArray();
                              using (Stream newFileStream = new MemoryStream(bytes))
                              {
                                  var byteBuffer = new byte[newFileStream.Length - 1];
                                  newFileStream.Read(byteBuffer, 0, byteBuffer.Length);
                                  zipToWrite.EnableZip64 = Zip64Option.Never;
                                  zipToWrite.AlternateEncodingUsage = ZipOption.Always;
                                  zipToWrite.AlternateEncoding = Encoding.GetEncoding(866);
                                  zipToWrite.PutNextEntry(nameFile);
                                  zipToWrite.Write(byteBuffer, 0, byteBuffer.Length);
                                  //Нужен ли здесь пароль на архивы
                              }
                          }
                        
                          i++;
                        }
                    }
                    zip.Close();
                }
                return File.ReadAllBytes(fullPathZip);
            }
            return null;
        }

        /// <summary>
        /// Генерация zip архива для отправки по SMTP
        /// </summary>
        /// <param name="collectionNameFile">Массив имен файлов</param>
        /// <param name="fullPathZip">Полный путь к наименованию Архива файла</param>
        /// <param name="isDoublicate"></param>
        /// <returns>Полный путь к архиву</returns>
        public byte[] StartZipArchiveOut(string[] collectionNameFile, string fullPathZip, bool isDoublicate = true)
        {
            try
            {
                if (collectionNameFile.Length > 0)
                {
                    using (var zip = File.Open(fullPathZip, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        using (var zipToWrite = new ZipOutputStream(zip))
                        {
                            var i = 0;
                            foreach (var fullNameFile in collectionNameFile)
                            {
                                var fileStream = File.ReadAllBytes(fullNameFile);
                                var nameFile = isDoublicate ? $"{i}_{Path.GetFileName(fullNameFile)}" : $"{Path.GetFileName(fullNameFile)}";
                                using (Stream newFileStream = new MemoryStream(fileStream))
                                {
                                    var byteBuffer = new byte[newFileStream.Length];
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
                    return File.ReadAllBytes(fullPathZip);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            return null;
        }
        /// <summary>
        /// Удаление всех файлов в папке 
        /// </summary>
        /// <param name="path">Путь к папке для удаления</param>
        public void DropAllFileToPath(string path)
        {
            foreach (FileInfo file in new DirectoryInfo(path).GetFiles())
            {
                Loggers.Log4NetLogger.Info(new Exception($"Наименование удаленных файлов: {file.FullName}"));
                file.Delete();
            }
        }
    }
}

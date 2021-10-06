using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using EfDatabase.Inventory.Base;
using EfDatabase.Inventory.MailLogicLotus;
using MimeKit;

namespace LibraryOutlook.SubscribeOutlook
{
   public class CalendarVks
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="collectionMessageAttach">Коллекция календарей</param>
        /// <param name="idMessage">Ун письма</param>
        public string CalendarParser(List<MimeEntity> collectionMessageAttach,string idMessage)
        {
            var mailCalendar = new MailLogicLotus();
            var calendar = new Calendar();
            var option = FormatOptions.Default.Clone();
            option.AllowMixedHeaderCharsets = false;
                foreach (var mimeEntity in collectionMessageAttach)
                {
                    try
                    {
                        string contentDecode;
                        using (var stream = new MemoryStream())
                        {
                            mimeEntity.WriteTo(option, stream);
                            contentDecode = Encoding.UTF8.GetString(stream.ToArray());
                        }
                        string[] separator = {Environment.NewLine};
                        contentDecode = string.Join(" ",contentDecode.Split(separator, StringSplitOptions.RemoveEmptyEntries));
                        var descriptionFull = Regex.Match((contentDecode), @"DESCRIPTION:(.+)DTEND").Value;
                        var descriptionAll = Regex.Match(descriptionFull, @"([^DESCRIPTION:](.+)[^DTEND])").Value;
                        var description = Regex.Match(descriptionAll, @"(.+=)").Value.TrimEnd(new[] { '=','И','Д', 'I','D','n','\\' });
                        var idVks = Regex.Match(descriptionFull, "([^=][0-9]{3,6})").Value;
                        var dateEndPresumably = Regex.Match(contentDecode, "DTEND;TZID=\"Russian Standard Time\":([0-9]+T[0-9]+Z?)").Value;
                        var dateStartPresumably = Regex.Match(contentDecode, "DTSTART;TZID=\"Russian Standard Time\":([0-9]+T[0-9]+Z?)").Value;
                        var dateEnd = Regex.Match(dateEndPresumably, "([0-9]+T[0-9]+Z?)").Value;
                        var dateStart = Regex.Match(dateStartPresumably, "([0-9]+T[0-9]+Z?)").Value;
                        calendar.DescriptionVKS = description;
                        calendar.FullDescription = descriptionAll;
                        calendar.IdMail = idMessage;
                        calendar.DateStart = Convert.ToDateTime(dateStart.Insert(4, "-").Insert(7, "-").Insert(13, ":").Insert(16, ":"));
                        calendar.DateFinish = Convert.ToDateTime(dateEnd.Insert(4, "-").Insert(7, "-").Insert(13, ":").Insert(16, ":"));
                        calendar.IdVKS = idVks;
                        mailCalendar.AddModelCalendar(calendar);
                    }
                    catch (Exception ex)
                    {
                        Loggers.Log4NetLogger.Error(new Exception($"Во время работы с календарем по письму {idMessage} произошла ошибка!"));
                        Loggers.Log4NetLogger.Error(ex);
                    }
                }
                mailCalendar.Dispose();
                return $"Добавлено в календарь запись о ВКС {calendar.FullDescription} ID {calendar.IdVKS} Дата начала {calendar.DateStart} Дата окончания {calendar.DateFinish}";
        }
        }
}

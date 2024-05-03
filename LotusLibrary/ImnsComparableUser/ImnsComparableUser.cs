using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Domino;
using LibaryXMLAuto.Inventarization.ModelComparableUserAllSystem;
using LotusLibrary.DbConnected;

namespace LotusLibrary.ImnsComparableUser
{
    public class ImnsComparableUser : IDisposable
    {
        ///Конфигурация Lotus
        private LotusConfig Config { get; set; }
        ///База данных Lotus
        private LotusConnectedDataBase Db { get; set; }

        /// <summary>
        /// Document Department
        /// </summary>
        private NotesDocument DocumentDepartment { get; set; }
        /// <summary>
        /// Document Users
        /// </summary>
        private NotesDocument DocumentUsers { get; set; }
        /// <summary>
        /// Коллекция документов Отделов
        /// </summary>
        private NotesDocumentCollection DocumentCollectionDepartment { get; set; }
        /// <summary>
        /// Коллекция документов Пользователей
        /// </summary>
        private NotesDocumentCollection DocumentCollectionUsers { get; set; }
        /// <summary>
        /// Поиск всех пользователей в Lotus 7751
        /// </summary>
        public ImnsComparableUser()
        {
            Dispose();
            Config = new LotusConfig();
            Db = new LotusConnectedDataBase(Config.LotusIdFilePassword);
        }

        /// <summary>
        /// Поиск всех пользователей в системе Lotus Notes
        /// </summary>
        /// <returns></returns>
        public ModelComparableUser FindAllUsersAndAttribute()
        {
            var modelComparableUsers = new ModelComparableUser();
            try
            {
                var modelList = new List<FullModelUserAllSystem>();
                Db.LotusConnectedDataBaseServer(Config.LotusServer, Config.LotusImns);
                DocumentCollectionDepartment = Db.Db.Search(String.Format("Select(Form= \"Department\")"), null, 0);
                DocumentDepartment = DocumentCollectionDepartment.GetFirstDocument();
                while (DocumentDepartment != null)
                {
                    string repDep = DocumentDepartment.GetItemValue("Abbreviation")[0];
                    DocumentCollectionUsers = Db.Db.Search(String.Format("Select (Abbreviation= \"" + repDep + "\"&Form= \"Department\"| @AllChildren)"), null, 0);
                    DocumentUsers = DocumentCollectionUsers.GetFirstDocument();
                    while (DocumentUsers != null)
                    { 
                        string nameNull = DocumentUsers.GetItemValue("Abbreviation")[0];
                        if (DocumentUsers.GetItemValue("Discharged")[0] != "1" && nameNull.Length == 0)
                        {
                            var modelUsers = new FullModelUserAllSystem();
                            modelUsers.SystemDataBase = "Lotus Notes";
                            modelUsers.Surname = DocumentUsers.GetItemValue("LastName")[0];
                            modelUsers.Name = DocumentUsers.GetItemValue("FirstName")[0];
                            modelUsers.Patronymic = DocumentUsers.GetItemValue("Person_MiddleName")[0];
                            modelUsers.FullName = DocumentUsers.GetItemValue("Person_Name")[0];
                            modelUsers.Department = repDep;
                            modelUsers.JobTitle = DocumentUsers.GetItemValue("Person_Post")[0];
                            modelUsers.Ranks = DocumentUsers.GetItemValue("Person_Title")[0];
                            modelUsers.Room = DocumentUsers.GetItemValue("Comm_Location")[0];
                            modelUsers.PhoneNumber = DocumentUsers.GetItemValue("Comm_Phone")[0]+"&"+DocumentUsers.GetItemValue("Comm_Ext")[0] ;
                            modelUsers.TypePhone = 1;
                            modelUsers.NameMail = DocumentUsers.GetItemValue("MailAddress")[0];
                            modelList.Add(modelUsers);
                        }
                        DocumentUsers = DocumentCollectionUsers.GetNextDocument(DocumentUsers);
                    }
                    DocumentDepartment = DocumentCollectionDepartment.GetNextDocument(DocumentDepartment);
                }
                modelComparableUsers.FullModelUserAllSystem = modelList;
            }
            catch (Exception ex)
            {
                Loggers.Log4NetLogger.Error(ex);
                throw;
            }
            finally
            {
                if (DocumentDepartment != null)
                    Marshal.ReleaseComObject(DocumentDepartment);
                DocumentDepartment = null;
                if (DocumentUsers != null)
                    Marshal.ReleaseComObject(DocumentUsers);
                DocumentUsers = null;
                if (DocumentCollectionDepartment != null)
                    Marshal.ReleaseComObject(DocumentCollectionDepartment);
                DocumentCollectionDepartment = null;
                if (DocumentCollectionUsers != null)
                    Marshal.ReleaseComObject(DocumentCollectionUsers);
                DocumentCollectionUsers = null;
                Loggers.Log4NetLogger.Info(new Exception("Сбор данных по справочнику ИМНС Завершен!!!"));
            }
            return modelComparableUsers;
        }
        /// <summary>
        /// Поиск ЗГ в лотус 8.5 по ИНН 662335054126
        /// </summary>
        /// <param name="inn">ИНН</param>
        public List<ModelFindZg.ModelFindZg> FindLotusNotesZg(string[] inn)
        {
            try
            {
                var modelZg = new List<ModelFindZg.ModelFindZg>();
      
                Db.LotusConnectedDataBaseServer(Config.LotusServer, "IFNS\\2012\\itof_zg_2012.nsf");
                foreach (var i in inn)
                {
                    var index = i.Split(' ');
                    DocumentCollectionUsers = Db.Db.Search(String.Format("@Select(@Contains(IO_INN;\"{0}\"))", index[0]), null, 0);
                    DocumentUsers = DocumentCollectionUsers.GetFirstDocument();
                    while (DocumentUsers != null)
                    {
                        modelZg.Add(new ModelFindZg.ModelFindZg()
                        {
                            FioFindMemo = i,
                            FioFindLotus = DocumentUsers.GetItemValue("InetFromName")[0],
                            Inn = DocumentUsers.GetItemValue("IO_INN")[0],
                            ZgNumber = DocumentUsers.GetItemValue("InCard_Index")[0],
                            InCard_RgDate = Convert.ToDateTime(DocumentUsers.GetItemValue("InCard_RgDate")[0]),
                            Ex_ExecDirect = DocumentUsers.GetItemValue("Ex_ExecDirect")[0],
                            Dept = DocumentUsers.GetItemValue("IOT_DEPT")[0],
                            OutNumber = DocumentUsers.GetItemValue("Ex_ExecutionMarks")[0],
                            CheckUpDate = string.IsNullOrWhiteSpace(DocumentUsers.GetItemValue("ctbCheckUpDate")?[0].ToString()) ? null : Convert.ToDateTime(DocumentUsers.GetItemValue("ctbCheckUpDate")[0]),
                            Number = DocumentUsers.GetItemValue("InCard_RespOutNum")[0],
                        });
                        DocumentUsers = DocumentCollectionUsers.GetNextDocument(DocumentUsers);
                    }
                }
                return modelZg;
            }
            catch (Exception ex)
            {
                Loggers.Log4NetLogger.Error(ex);
                throw;
            }
            finally
            {
                if (DocumentDepartment != null)
                    Marshal.ReleaseComObject(DocumentDepartment);
                DocumentDepartment = null;
                if (DocumentUsers != null)
                    Marshal.ReleaseComObject(DocumentUsers);
                DocumentUsers = null;
                if (DocumentCollectionDepartment != null)
                    Marshal.ReleaseComObject(DocumentCollectionDepartment);
                DocumentCollectionDepartment = null;
                if (DocumentCollectionUsers != null)
                    Marshal.ReleaseComObject(DocumentCollectionUsers);
                DocumentCollectionUsers = null;
                Loggers.Log4NetLogger.Info(new Exception("Сбор данных по справочнику ИМНС Завершен!!!"));
            }
        }


        private void ReleaseUnmanagedResources()
        {

        }

        private void Dispose(bool disposing)
        {
            ReleaseUnmanagedResources();
            if (disposing)
            {
                Db?.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
              //GC.SuppressFinalize(this);
        }

        ~ImnsComparableUser()
        {
            Dispose(false);
        }

    }
}

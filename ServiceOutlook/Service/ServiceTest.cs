using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibaryXMLAuto.Inventarization.ModelComparableUserAllSystem;
using LotusLibrary.ImnsComparableUser;

namespace ServiceOutlook.Service
{
    class ServiceTest : IServiceTest
    {
        public string GenerateSqlSelect()
        {
            return "Вроде запустили";
        }
        /// <summary>
        /// Все пользователи Lotus Notes
        /// </summary>
        /// <returns></returns>
        public async Task<ModelComparableUser> AllUsersLotusNotes()
        {
             return await Task.Factory.StartNew(() =>
            {
                ImnsComparableUser modelImns = new ImnsComparableUser();
                var model = modelImns.FindAllUsersAndAttribute();
                modelImns.Dispose();
                return model;
            });
        }
    }
}

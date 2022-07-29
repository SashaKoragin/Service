using LibaryXMLAuto.Inventarization.ModelComparableUserAllSystem;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Threading.Tasks;

namespace ServiceOutlook.Service
{
    [ServiceContract]
    interface IServiceTest
    {
        /// <summary>
        /// http://77068-app065:8585/ServiceOutlook/Test
        /// Генерация запросов на клиент
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        [WebInvoke(Method = "GET", RequestFormat = WebMessageFormat.Json, UriTemplate = "/Test", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        string GenerateSqlSelect();

        /// <summary>
        /// http://77068-app065:8585/ServiceOutlook/AllUsersLotusNotes
        /// Генерация запросов на клиент
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        [WebInvoke(Method = "GET", RequestFormat = WebMessageFormat.Json, UriTemplate = "/AllUsersLotusNotes",
            ResponseFormat = WebMessageFormat.Xml, BodyStyle = WebMessageBodyStyle.Bare)]
       Task<ModelComparableUser> AllUsersLotusNotes();
    }
}

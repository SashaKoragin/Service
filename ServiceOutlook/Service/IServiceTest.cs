using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Threading.Tasks;

namespace ServiceOutlook.Service
{
    [ServiceContract]
    interface IServiceTest
    {
        /// <summary>
        /// http://localhost:8182/Inventarka/GenerateSqlSelect
        /// Генерация запросов на клиент
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        [WebInvoke(Method = "GET", RequestFormat = WebMessageFormat.Json, UriTemplate = "/Test", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        string GenerateSqlSelect();

    }
}

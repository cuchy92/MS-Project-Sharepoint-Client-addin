using System;
using System.Net;

namespace ProjectSharePointSync.Core.SharePoint
{
    public sealed class GraphException : Exception
    {
        public GraphException(HttpStatusCode status, string code, string message)
            : base(string.Format("Microsoft Graph error {0} ({1}): {2}", (int)status, code, message))
        {
            Status = status;
            Code = code;
        }

        public HttpStatusCode Status { get; private set; }

        public string Code { get; private set; }
    }
}

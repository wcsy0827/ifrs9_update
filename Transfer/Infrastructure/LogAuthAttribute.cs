using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Transfer.Controllers;
using Transfer.Models;
using Transfer.Utility;

namespace Transfer.Infrastructure
{
    public class LogAuthAttribute : AuthorizeAttribute
    {
        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            bool hasuser = httpContext.User != null;
            bool isAuthenticated = hasuser && httpContext.User.Identity.IsAuthenticated;
            return isAuthenticated;
        }

        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            base.OnAuthorization(filterContext);
        }

        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            if (filterContext.HttpContext.Request.IsAjaxRequest())
            {
                filterContext.Result = new ContentResult();
                filterContext.HttpContext.Response.StatusCode = 401;
            }
            else
            {
                filterContext.Result = new HttpUnauthorizedResult();
            }          
        }
    }
}
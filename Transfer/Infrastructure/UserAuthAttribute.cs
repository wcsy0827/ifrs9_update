using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Transfer.Controllers;
using Transfer.Models;
using Transfer.Utility;

namespace Transfer.Infrastructure
{
    public class UserAuthAttribute : AuthorizeAttribute
    {
        private string _href;

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            string controller = (string)((MvcHandler)httpContext.Handler).RequestContext.RouteData.Values["controller"];
            string action = (string)((MvcHandler)httpContext.Handler).RequestContext.RouteData.Values["action"];
            _href = string.Format("{0},{1}", action, controller);
            bool flag = false;
            string menu_id = string.Empty;
            bool systemFlag = false;
            if (controller == "Home" && action == "Index")
                return true;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var sub = db.IFRS9_Menu_Sub.FirstOrDefault(x => x.Href != null && 
                x.Href.Replace(System.Environment.NewLine,string.Empty).Trim().Equals(_href));
                if (sub != null) //action 為view
                {
                    menu_id = sub.Menu_Id;
                    systemFlag = sub.Menu == "System";
                }
                else //action 為event
                {
                    return true;
                }
                if (systemFlag)
                {
                    flag = db.IFRS9_User.FirstOrDefault(x =>
                    x.User_Account == AccountController.CurrentUserInfo.Name &&
                    x.AdminFlag) != null;
                }
                else
                {
                    flag = db.IFRS9_Menu_Set.FirstOrDefault(
                         x => x.User_Account.Equals(AccountController.CurrentUserInfo.Name) &&
                         x.Menu_Id.Equals(menu_id) && x.Effective == "Y") != null;
                }
                if (flag)
                {
                    var dt = DateTime.Now.Date;
                    var lt = TypeTransfer.stringToDateTime(AccountController.CurrentUserInfo.Ticket.UserData);
                    var userlog = db.IFRS9_User_Log.AsNoTracking()
                        .Where(x => x.User_Account == AccountController.CurrentUserInfo.Name &&
                        x.Login_Date == dt)
                        .AsEnumerable()
                        .Where(x=>x.Login_Time.ToString("yyyy/MM/dd HH:mm:ss") == AccountController.CurrentUserInfo.Ticket.UserData).FirstOrDefault();
                    if (userlog != null)
                    db.IFRS9_Browse_Log.Add(new IFRS9_Browse_Log()
                    {
                        User_Account = AccountController.CurrentUserInfo.Name,
                        Browse_Time = DateTime.Now,
                        Menu_Id = menu_id,
                        Login_Time = userlog.Login_Time
                    });
                    try
                    {
                        db.SaveChanges();
                    }
                    catch(Exception ex)
                    {
                        ex.exceptionMessage();
                    }
                }
            }
            return flag;
        }

        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            filterContext.Result = new RedirectResult("/Home/Error401");
        }
    }
}
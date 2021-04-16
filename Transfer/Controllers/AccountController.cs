﻿using BotDetect.Web.Mvc;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using System.Web.Security;
using Transfer.Infrastructure;
using Transfer.Models;
using Transfer.Utility;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    public class AccountController : Controller
    {
        public static FormsIdentity CurrentUserInfo
        {
            get
            {
                var httpContext = System.Web.HttpContext.Current;
                var identity = httpContext.User.Identity.IsAuthenticated;
                if (identity)
                {
                    FormsIdentity id = (FormsIdentity)httpContext.User.Identity;
                    return id;
                }
                else
                    return null;
            }
        }

        public static string CurrentUserName
        {
            get
            {
                var httpContext = System.Web.HttpContext.Current;
                var identity = httpContext.User.Identity.IsAuthenticated;
                if (identity)
                {
                    string name = string.Empty;
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        var user = db.IFRS9_User.AsNoTracking()
                              .FirstOrDefault(x => x.User_Account == httpContext.User.Identity.Name);
                        if (user != null)
                            name = user.User_Name;
                    }
                    return name;
                }
                return string.Empty;
            }
        }

        // GET: Account
        [AllowAnonymous]
        [ValidateInput(false)]
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [ValidateInput(false)]
        //[CaptchaValidation("CaptchaCode", "ExampleCaptcha", "驗證碼不正確!")]
        public ActionResult Logon(string userId, string pwd)
        {
                Session.Clear();
                Session.Abandon();
                Session.RemoveAll();
                if (Request.Cookies["ASP.NET_SessionId"] != null)
                {
                    Response.Cookies["ASP.NET_SessionId"].Value = string.Empty;
                    Response.Cookies["ASP.NET_SessionId"].Expires = DateTime.Now.AddMonths(-20);
                }
                if (Request.Cookies["adAuthCookie"] != null)
                {
                    Response.Cookies["adAuthCookie"].Value = string.Empty;
                    Response.Cookies["adAuthCookie"].Expires = DateTime.Now.AddMonths(-20);
                }
                if (Request.Cookies["__RequestVerificationToken"] != null)
                {
                    Response.Cookies["__RequestVerificationToken"].Value = string.Empty;
                    Response.Cookies["__RequestVerificationToken"].Expires = DateTime.Now.AddMonths(-20);
                }

                bool flag = false;
                var now = DateTime.Now;
            #region txtlog 檔案名稱 
            //191113 John.解金檢議題，加入Hostname相關
            //加入儲存登入資訊txt檔機制與路徑
            string Logintxtpath = $"{DateTime.Now.ToString("yyyyMMdd")}_{SetFile.LoginLog}"; //預設txt名稱
            string LoginExcptiontxtpath = $"{DateTime.Now.ToString("yyyyMMdd")}_{SetFile.LoginExceptionLog}"; //預設txt名稱
            string LoginDetailInfo = string.Empty;
            #endregion txtlog 檔案名稱
            //if (!ModelState.IsValid)
            //{
            //    TempData["User"] = userId;
            //    TempData["Login"] = Message_Type.login_Captcha_Fail.GetDescription();
            //    return RedirectToAction("Login", "Account");
            //}
            //else
            //{
            //MvcCaptcha.ResetCaptcha("ExampleCaptcha");
            FileRelated.createFile(@"D:\IFRS9Log");
                try
                {
                    // set up domain context
                    //PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                    // find the current user
                    //UserPrincipal aduser = UserPrincipal.Current;

                    //驗證AD帳號 
                    flag = LdapAuthentication.isAuthenticatrd(userId, pwd);
                }
                catch(Exception ex)
                {
                    TxtLog.detailtxtLog(false,ex.ToString(), txtLocation(LoginExcptiontxtpath));
                }
                var user = new IFRS9_User();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (flag) //AD帳號
                    {
                        user = db.IFRS9_User.AsNoTracking().FirstOrDefault(x => x.User_Account == userId);
                    }
                    else //非AD帳號
                    {
                        user = db.IFRS9_User
                            .Where(x => userId.Equals(x.User_Account))
                            .AsEnumerable()
                            .FirstOrDefault(x => pwd.stringToSHA512().Equals(x.User_Password));
                    }
                    if (user != null)
                    {
                        if (user.Effective)
                        {
                            this.LoginProcess(user.User_Account, false, now);
                            flag = true;
                            user.LoginFlag = true;
                            string IP = System.Web.HttpContext.Current.Request
                                .ServerVariables["HTTP_X_FORWARDED_FOR"];
                            if (!string.IsNullOrEmpty(IP))
                            {
                                string[] addresses = IP.Split(',');
                                if (addresses.Length != 0)
                                {
                                    IP= addresses[0];
                                }
                            }
                            else
                            {
                                IP= System.Web.HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
                            }

                            //1002 John.解WebInpsect弱點修正_抓hostname
                            try
                            {
                                IPHostEntry ipHostName = Dns.GetHostEntry(IP?.Trim() ?? string.Empty);
                                var IPAddrMachine = ipHostName.HostName.ToString();
                                IP = IP + ";Machine:" + IPAddrMachine;

                            }
                            catch(Exception ex)
                            {
                            IP = IP + ";機器名稱不明";
                            TxtLog.detailtxtLog(false, ex.ToString(), txtLocation(LoginExcptiontxtpath));
                        }
                            db.IFRS9_User_Log.Add(
                                new IFRS9_User_Log()
                                {
                                    User_Account = user.User_Account,
                                    Login_Time = now,
                                    Login_IP = IP,
                                    Login_Date = now.Date
                                });
                            try
                            {
                                db.SaveChanges();
                                TxtLog.detailtxtLog(true,string.Format("{0}{1}", $"[{user.User_Account}]", $"[{IP}]"), txtLocation(Logintxtpath));
                            }
                            catch(Exception ex)
                            {
                                TxtLog.detailtxtLog(false,ex.ToString(), txtLocation(LoginExcptiontxtpath));
                            }
                        }
                        else
                        {
                            TempData["Login"] = Message_Type.login_Effective_Fail
                                .GetDescription();
                            ModelState.AddModelError("", Message_Type.login_Effective_Fail
                                .GetDescription());
                            flag = false;
                        }
                    }
                    else
                    {
                        TempData["Login"] = Message_Type.login_Fail.GetDescription();
                        ModelState.AddModelError("", Message_Type.login_Fail.GetDescription());
                        flag = false;
                    }
                }
                if (flag)
                {
                    TempData["Login"] = string.Empty;
                    return RedirectToAction("Index", "Home");
                }
                else
                {
                    TempData["User"] = userId;
                    return RedirectToAction("Login", "Account");
            }
                

            //}
        }

        [AllowAnonymous]
        [HttpGet]
        public ActionResult Logout()
        {

            //reset Captcha
            //MvcCaptcha.ResetCaptcha("ExampleCaptcha");

            bool hasuser = System.Web.HttpContext.Current.User != null;
            bool isAuthenticated = hasuser && System.Web.HttpContext.Current.User.Identity.IsAuthenticated;
            if (isAuthenticated)
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    string Login_Time = CurrentUserInfo.Ticket.UserData;
                    var dt = DateTime.Now.Date;
                    var user = db.IFRS9_User.FirstOrDefault(x => x.User_Account == CurrentUserInfo.Name);
                    var userLog = db.IFRS9_User_Log.Where(
                        x => x.User_Account == CurrentUserInfo.Name &&
                        x.Login_Date == dt)
                        .AsEnumerable()
                        .FirstOrDefault(x => x.Login_Time.ToString("yyyy/MM/dd HH:mm:ss") == Login_Time);
                    try
                    {
                        if (user != null && userLog != null)
                        {

                            string sql = $@"
update IFRS9_User
set LoginFlag = 'false'
where User_Account = {user.User_Account.stringToStrSql()}

update IFRS9_User_Log
set Logout_Time = '{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'
where  User_Account = {user.User_Account.stringToStrSql()}
and FORMAT(Login_Time, N'yyyy/MM/dd HH:mm:ss')   = '{Login_Time}'
";
                            db.Database.ExecuteSqlCommand(sql);
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.exceptionMessage();
                    }
                }

            }
            //清除所有的 session
            Session.Clear();
            Session.Abandon();
            Session.RemoveAll();
            if (Request.Cookies["ASP.NET_SessionId"] != null)
            {
                Response.Cookies["ASP.NET_SessionId"].Value = string.Empty;
                Response.Cookies["ASP.NET_SessionId"].Expires = DateTime.Now.AddMonths(-20);
            }
            if (Request.Cookies["adAuthCookie"] != null)
            {
                Response.Cookies["adAuthCookie"].Value = string.Empty;
                Response.Cookies["adAuthCookie"].Expires = DateTime.Now.AddMonths(-20);
            }
            if (Request.Cookies["__RequestVerificationToken"] != null)
            {
                Response.Cookies["__RequestVerificationToken"].Value = string.Empty;
                Response.Cookies["__RequestVerificationToken"].Expires = DateTime.Now.AddMonths(-20);
            }


            ////建立一個同名的 Cookie 來覆蓋原本的 Cookie
            //HttpCookie cookie1 = new HttpCookie(FormsAuthentication.FormsCookieName, "");
            //cookie1.Expires = DateTime.Now.AddYears(-1);
            //Response.Cookies.Add(cookie1);

            ////建立 ASP.NET 的 Session Cookie 同樣是為了覆蓋
            //HttpCookie cookie2 = new HttpCookie("ASP.NET_SessionId", "");
            //cookie2.Expires = DateTime.Now.AddYears(-1);
            //Response.Cookies.Add(cookie2);

            //登出
            FormsAuthentication.SignOut();

            TempData["Logout"] = "true";
            return RedirectToAction("Login");
        }

        [ChildActionOnly]
        public ActionResult Menu()
        {
            bool hasuser = System.Web.HttpContext.Current.User != null;
            bool isAuthenticated = hasuser && System.Web.HttpContext.Current.User.Identity.IsAuthenticated;

            if (isAuthenticated)
            {
                List<IFRS9_Menu_Sub> subs = new List<IFRS9_Menu_Sub>();
                List<IFRS9_Menu_Main> mains = new List<IFRS9_Menu_Main>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var _user = db.IFRS9_User.Where(x => x.User_Account == AccountController.CurrentUserInfo.Name && x.Effective).ToList();

                    if (_user.Any(x =>x.AdminFlag))
                    {
                        subs = db.IFRS9_Menu_Sub.AsNoTracking().Where(x => x.Menu == "System").OrderBy(x => x.Menu_Id).ToList();
                        mains = db.IFRS9_Menu_Main.AsNoTracking().Where(x => x.Menu == "System").ToList();
                    }
                    else if(_user.Any())
                    {
                        var _DebtType = _user.First().DebtType;
                        subs = db.IFRS9_Menu_Set.AsNoTracking()
                         .Where(x => x.User_Account == AccountController.CurrentUserInfo.Name && x.Effective == "Y")
                         .Select(x => x.IFRS9_Menu_Sub)
                         .Where(x => x.DebtType == "A" || x.DebtType == _DebtType, _DebtType != "A")
                         .OrderBy(x => x.Menu_Id).ToList();
                        var Menus = subs.Select(x => x.Menu).Distinct();
                        mains = db.IFRS9_Menu_Main.AsNoTracking()
                            .Where(x => Menus.Contains(x.Menu))
                            .OrderBy(x => x.Menu).ToList();
                    }
                }
                MenuModel menus = new MenuModel()
                {
                    menu_Main = mains,
                    menu_Sub = subs
                };

                return PartialView(menus);
            }
            else
            {
                return RedirectToAction("Login", "Account");
            }
        }
        #region PrivateFunction
        private void LoginProcess(string user, bool isRemeber, DateTime dt)
        {
            var ticket = new FormsAuthenticationTicket(
                version: 1,
                name: user,
                issueDate: dt,
                expiration: dt.AddMinutes(30),
                isPersistent: isRemeber,
                //userData: user,
                userData: dt.ToString("yyyy/MM/dd HH:mm:ss"),
                cookiePath: FormsAuthentication.FormsCookiePath);

            var encryptedTicket = FormsAuthentication.Encrypt(ticket);
            var cookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
            Response.Cookies.Add(cookie);
        }
        protected string txtLocation(string path) //191113 John.解金檢議題，加入Hostname相關
        {
            try
            {
                string projectFile = Server.MapPath("~/" + SetFile.IFRS9Log); //預設放Log的資料夾路徑
                string configTxtLocation = ConfigurationManager.AppSettings["txtLogLocation"];
                if (!string.IsNullOrWhiteSpace(configTxtLocation))
                    projectFile = configTxtLocation; //有設定webConfig且不為空就取代
                FileRelated.createFile(projectFile);
                string folderPath = Path.Combine(projectFile, path); //合併路徑&檔名
                return folderPath;
            }
            catch
            {
                return string.Empty;
            }
        }
        #endregion
    }

    public class MenuModel
    {
        public List<IFRS9_Menu_Main> menu_Main { get; set; }
        public List<IFRS9_Menu_Sub> menu_Sub { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Text.RegularExpressions;

namespace ExcelWebAppRedirecter
{
    /// <summary>
    /// Some web articles that bringed additional info for making this module.
    /// https://msdn.microsoft.com/en-us/library/ms474633.aspx
    /// http://stackoverflow.com/questions/2554044/open-an-spweb-from-a-single-url
    /// http://stackoverflow.com/questions/11633241/redirect-url-using-httpmodule-asp-net
    /// http://stackoverflow.com/questions/679559/debugging-an-httpmodule-on-the-asp-net-development-server
    /// https://parago.net/2011/01/16/how-to-implement-a-custom-sharepoint-2010-logging-service-for-uls-and-windows-event-log/
    /// 
    /// One of these entries must be added to the web.config in order to consume this HTTP module.
    /// <httpModules>
    ///     <add name = "ExcelWebAppRedirecter" type="ExcelWebAppRedirecter.Redirecter, ExcelWebAppRedirecter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=ab1b281a9b9b52f9" />
    /// </httpModules>
    /// <modules runAllManagedModulesForAllRequests = "true" >
    ///     <add name="ExcelWebAppRedirecter" type="ExcelWebAppRedirecter.Redirecter, ExcelWebAppRedirecter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=ab1b281a9b9b52f9" />
    /// </modules>
    /// 
    /// In development you can use gacutil for a faster deployment.
    /// C:\Program Files(x86)\Microsoft Visual Studio 14.0>gacutil -i "C:\Users\Administrator\Documents\visual studio 2015\Projects\ExcelWebAppRedirecter\ExcelWebAppRedirecter\bin\Debug\ExcelWebAppRedirecter.dll"
    /// C:\Program Files(x86)\Microsoft Visual Studio 14.0>gacutil -uf "ExcelWebAppRedirecter"
    /// </summary>
    public class Redirecter : IHttpModule
    {
        public void Dispose()
        {
            //throw new NotImplementedException();
        }

        public void Init(HttpApplication context)
        {
            try
            {
                context.BeginRequest += Context_BeginRequest;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't register the BeginRequest event! 'Excel Web App Redirecter' feature will not provide additional functionality from now on. Exception Message: " + ex.Message + ". Exception StackTrace: " + ex.StackTrace);
            }
        }

        private void Context_BeginRequest(object sender, EventArgs e)
        {
            try
            {
                HttpContext context = ((HttpApplication)sender).Context;
                if (context.Request.Url.AbsolutePath.EndsWith("xlviewer.aspx") && context.Request.QueryString.HasKeys() && context.Request.QueryString.AllKeys.Contains("id"))
                {
                    System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> An Excel Services HTTP request was received. The 'Excel Web App Redirecter' feature will redirect the request to Excel Web App.");
                    System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> The initial Excel Services HTTP request is: " + context.Request.Url.ToString() + ".");
                    System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> The excel file SharePoint path is: " + context.Request.QueryString.Get("id") + ". This value will be provided later on as 'sourcedoc' GET parameter.");

                    String strUrlToExtractBaseFrom = String.Empty;
                    Regex rgxExtractBaseURL = new Regex(@"^(http://|https://)([^/\s]+)");
                    MatchCollection mcMatches = rgxExtractBaseURL.Matches(context.Request.Url.ToString());
                    if(mcMatches!=null && mcMatches.Count==1)
                    {
                        strUrlToExtractBaseFrom = mcMatches[0].Value + context.Request.QueryString.Get("id");
                        System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> The URL base on which I will extract the base SharePoint Web URL is: " + strUrlToExtractBaseFrom + ".");
                        try
                        {
                            String strRedirectURL = String.Empty;
                            using (SPSite site = new SPSite(strUrlToExtractBaseFrom))
                            {
                                try
                                {
                                    //this ensures the URL is assigned to an existent SharePoint Site Collection.
                                    using (SPWeb web = site.OpenWeb())
                                    {
                                        //this ensures the URL is assigned to an existent SharePoint Web and the Web can be accessed.
                                        System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> The base SharePoint Web URL for the request is: " + web.Url + ".");
                                        strRedirectURL += web.Url.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't validate URL in the context of SharePoint Webs. Exception Message: " + ex.Message + ". Exception StackTrace: " + ex.StackTrace);
                                }
                            }


                            if (strRedirectURL != String.Empty)
                            {
                                strRedirectURL += "/_layouts/15/WopiFrame.aspx?sourcedoc=" + context.Request.QueryString.Get("id") + "&action=default";
                                System.Diagnostics.Trace.TraceInformation("'Excel Web App Redirecter' feature -> The redirect URL is: " + strRedirectURL + ". Enjoy Excel Web App instead of Excel Services!");

                                //do the redirect
                                context.Response.Redirect(strRedirectURL);
                            }
                            else
                            {
                                System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't construct a valid Excel Web App URL based on this Http request.");
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!(ex is System.Threading.ThreadAbortException))
                            {
                                System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't validate URL in the context of SharePoint site collections, or couldn't construct a valid Excel Web App URL. Exception Message: " + ex.Message + ". Exception StackTrace: " + ex.StackTrace);
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't construct a valid Excel Web App URL based on this Http request.");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!(ex is System.Threading.ThreadAbortException))
                {
                    System.Diagnostics.Trace.TraceError("'Excel Web App Redirecter' feature -> Couldn't analyse the HTTP request. Exception Message: " + ex.Message + ". Exception StackTrace: " + ex.StackTrace);
                }
            }
        }
    }
}

using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using System.DirectoryServices;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;
using System.Text;
using System.Data;

namespace Bot_Application1 {
    [BotAuthentication]
    public class MessagesController : ApiController {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity) {
            if (activity.Type == ActivityTypes.Message) {
                ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));
                // calculate something for us to return
                int length = (activity.Text ?? string.Empty).Length;
                string replyBot = string.Empty;
                string emailid = string.Empty;
                if (activity.Text.TrimEnd().ToString().Contains("@hpe.com")) {
                    emailid = activity.Text.Substring(activity.Text.LastIndexOf(" ")).TrimStart();
                    replyBot = GetEmpDetails(emailid);
                } else if (activity.Text.ToString().ToLower().Contains("thank")) {
                    replyBot = "You are welcome";
                } else if (activity.Text.ToString().ToLower().Contains("hi")) {
                    replyBot = "hi, how can i help you?";
                } else {
                    replyBot = "I dont understand you. I only understand email ids ";
                }
                // return our reply to the user
                Activity reply = activity.CreateReply(replyBot.ToLower().Contains("exception") ? "An error occurred while processing: " + replyBot : replyBot);
                await connector.Conversations.ReplyToActivityAsync(reply);
            } else {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        private void SendEmail() {
            throw new NotImplementedException();
        }
        private static DirectorySearcher SearchLDAP(string strUser) {
            DirectorySearcher searcher;
            StringBuilder builder = new StringBuilder();
            try {
                string strServerDNS = "hpe-pro-ods-ed.infra.hpecorp.net";
                string strSearchBaseDN = "ou=People,o=hp.com";
                builder = new StringBuilder();
                builder.Clear();
                builder.Append("LDAP://" + strServerDNS + "/" + strSearchBaseDN);
                string strLDAPPath = builder.ToString();
                DirectoryEntry objDirEntry = new DirectoryEntry();
                objDirEntry.AuthenticationType = AuthenticationTypes.None;
                objDirEntry.Path = strLDAPPath;
                searcher = new DirectorySearcher(objDirEntry);
                searcher.SearchRoot = objDirEntry;
                searcher.SearchScope = SearchScope.Subtree;
                if (strUser.Length > 0) {
                    if (!strUser.Contains("@"))
                        searcher.Filter = "(ntUserDomainId=" + strUser + ")";
                    else
                        searcher.Filter = "(uid=" + strUser + ")";
                }
                return searcher;
            } catch { return null; } finally { builder.Clear(); }
        }

        public string GetEmpDetails(string strUser) {
            string FUNC_NAME =
               System.Reflection.MethodBase.GetCurrentMethod().Name + "(): ";

            //strUser = strUser.Replace("`", "@");
            // string strUserEmail = "soumarshi-nag%1biswas%2hp%1com";
            string userName = string.Empty;
            string managerName = string.Empty;
            string progress = "begin admin service GetEmpDetails method.";
            DirectorySearcher searcher = SearchLDAP(strUser);
            if (searcher != null) {
                searcher.PropertiesToLoad.Add("givenName");
                searcher.PropertiesToLoad.Add("hpaltlegalname");
                searcher.PropertiesToLoad.Add("hppicturethumbnailuri");
                searcher.PropertiesToLoad.Add("sn");
                searcher.PropertiesToLoad.Add("uid");
                searcher.PropertiesToLoad.Add("ntuserdomainid");
                searcher.PropertiesToLoad.Add("mobile");
                searcher.PropertiesToLoad.Add("l");
                searcher.PropertiesToLoad.Add("manager");
                searcher.PropertiesToLoad.Add("buildingName");
            }
            progress = "Method to Get the Employee Details";
            SearchResultCollection resultcollection = null;
            try {
                resultcollection = searcher.FindAll();
                progress = "Datatable to Get the User Details";
                DataTable dtresult = new DataTable("Results");
                foreach (string colName in searcher.PropertiesToLoad) {
                    if (colName == "l")
                        dtresult.Columns.Add("locality", System.Type.GetType("System.String"));
                    else
                        dtresult.Columns.Add(colName, System.Type.GetType("System.String"));
                }

                string strtemp = "";

                foreach (SearchResult objresult in resultcollection) {
                    DataRow dr = dtresult.NewRow();
                    foreach (string colName in searcher.PropertiesToLoad) {
                        if (colName == "l") {
                            strtemp = objresult.Properties["l"][0].ToString();
                            dr["locality"] = strtemp;
                        } else if (objresult.Properties.Contains(colName)) {
                            strtemp = objresult.Properties[colName][0].ToString();
                            strtemp = strtemp.Replace("$", "");
                            if (colName == "ntuserdomainid") {
                                strtemp = strtemp.Replace(":", "\\");
                            }
                            dr[colName] = strtemp;
                        } else {
                            dr[colName] = "";
                        }
                    }
                    dtresult.Columns.Remove("ADSPath");
                    dtresult.Rows.Add(dr);
                }
                searcher = SearchLDAP(dtresult.Rows[0]["manager"].ToString().Substring(4).Substring(0, dtresult.Rows[0]["manager"].ToString().Substring(4).IndexOf(",")));
                if (searcher != null) {
                    searcher.PropertiesToLoad.Add("ntuserdomainid");
                    searcher.PropertiesToLoad.Add("givenName");
                    searcher.PropertiesToLoad.Add("sn");
                }
                progress = "Method to Get the Employee Details";
                resultcollection = null;
                try {
                    resultcollection = searcher.FindAll();
                } catch (Exception ex) {
                    throw;
                }
                foreach (SearchResult objresult in resultcollection) {
                    //strtemp = objresult.Properties["ntuserdomainid"][0].ToString();
                    strtemp = objresult.Properties["givenName"][0].ToString() + " " + objresult.Properties["sn"][0].ToString();
                    dtresult.Rows[0]["manager"] = strtemp.Replace(":", "\\");
                    userName = dtresult.Rows[0]["givenName"].ToString() + " " + dtresult.Rows[0]["sn"].ToString();
                    managerName = dtresult.Rows[0]["manager"].ToString();
                    // string uri = dtresult.Rows[0]["hppicturethumbnailuri"].ToString();
                }
            } catch (Exception ex) {
                return ex.Message.ToString();
            }
            return $"The Given Name of {strUser} is {userName} and manager name is {managerName}";
            // return JSONHelper(dtresult, "Found data successfully").ToString();
        }
        private Activity HandleSystemMessage(Activity message) {
            if (message.Type == ActivityTypes.DeleteUserData) {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            } else if (message.Type == ActivityTypes.ConversationUpdate) {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            } else if (message.Type == ActivityTypes.ContactRelationUpdate) {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            } else if (message.Type == ActivityTypes.Typing) {
                // Handle knowing tha the user is typing
            } else if (message.Type == ActivityTypes.Ping) {
            }

            return null;
        }
    }
}
using CamlexNET;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Configuration;
using System.Collections.Specialized;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils.Helpers;
using Utils.SPUtil;

namespace ChangeLink
{
    class Program
    {
        public static class Constantes
        {
            public static readonly string URL = ConfigurationManager.AppSettings.Get("URL");
            public static readonly string ColumnName = ConfigurationManager.AppSettings.Get("ColumnName");
        }
        static void Main(string[] args)
        {
            using (var cc = GetContextMFA())
            {
                int count = 0;
                var listCollection = cc.Web.Lists;
                cc.Load(listCollection);
                cc.Load(cc.Web);
                cc.ExecuteQuery();
                FieldUrlValue url = new FieldUrlValue();
                url.Url = cc.Web.Url;

                IQueryable<List> lsSPLists = listCollection.Where(x => x.Title.ToLower().Contains("historico"));
                foreach (List oList in lsSPLists)
                {
                    var cml = Camlex.Query()
                        .ViewFields(x => new[] { x["ID"], x[Constantes.ColumnName] })
                        .Take(4999)
                        .OrderBy(x => (x["ID"]))
                        .ToCamlQuery();
                    var itemsCol = ListUtil.GetAllItemsFrom(cc, "/Lists/" + oList.Title, cml, false);

                    itemsCol.RemoveAll(x => {
                        var u = (FieldUrlValue)x.FieldValues[Constantes.ColumnName];
                        return u.Url.Contains("wikidocumentos");
                    });
                    foreach (ListItem lsItem in itemsCol)
                    {
                        lsItem[Constantes.ColumnName] = url;
                        lsItem.SystemUpdate();
                        if (count == 50)
                        {
                            cc.ExecuteQuery();
                            count = 0;
                        }
                        count++;
                    }
                    cc.ExecuteQuery();
                }
            }
        }
        static  ClientContext GetContextMFA()
        {
            var authManager = new AuthenticationManager();
            return authManager.GetWebLoginClientContext(Constantes.URL);
        }
    }
}

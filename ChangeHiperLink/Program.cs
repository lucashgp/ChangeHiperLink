using CamlexNET;
using System;
using Utils.Helpers;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace ChangeHiperLink
{
    class Program
    {
        public static class Constantes
        {
            public static readonly string URL = @"https://lhgpms.sharepoint.com/";
            public static readonly string Login = "lucas.padilha@lhgpms.onmicrosoft.com";
            public static readonly string Senha = "Vaicharmander1!";
            public static readonly string ClientId = "65bb24aa-068f-4e37-a71a-d7dec2a3876c";
            public static readonly string ClientSecret = "4uraIjs2mph1YePV6sU9S2K5J7x3pXvf+psiRXyVaNw=";
        }
        static void Main(string[] args)
        {
            ChangeLinks();
        }
        static void ChangeLinks()
        {
            using (var cc = ContextHelper.CreateAppOnlyClientContext(Constantes.URL, Constantes.ClientId, Constantes.ClientSecret))
            {
                var lists = cc.Web.Lists;
                cc.Load(lists);
                cc.ExecuteQuery();
                var lista = lists.GetByTitle("Historico1");
                cc.Load(lists, x => x.Where(l => l.Title.StartsWith("Historico")));

                var cml = Camlex.Query()
                    .ViewFields(x => new[] { x["ID"], x["Url"] })
                    .Take(4999)
                    .Where(x => !(x["Url"].ToString().Contains("wikidocumentos")))
                    .OrderBy(x => (x["ID"]))
                    .ToCamlQuery();
                var itemsCol = lista.GetItems(cml);

                foreach (var lsItem in itemsCol)
                {
                    lsItem["Url"] = cc.Web;
                }


            }
        }
    }
}

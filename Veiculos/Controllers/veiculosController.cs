using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Mvc;
using Veiculos.Models;

namespace Veiculos.Controllers
{
    public class VeiculosController : ApiController
    {
        private vdlwebcontrolEntities db = new vdlwebcontrolEntities();

        class VehicleResults
        {
            public int Year { get; set; }
            public int Month { get; set; }
            public double Amount { get; set; }
        }

        public string RemoveAcentuation(string text)
        {
            return System.Web.HttpUtility.UrlDecode(System.Web.HttpUtility.UrlEncode(text, System.Text.Encoding.GetEncoding("iso-8859-7")));
        }

        [System.Web.Http.HttpGet]
        public IHttpActionResult TypesAutocomplete(string Prefix)
        {
            List<t_ecf_vehicle_types> ObjList = db.t_ecf_vehicle_types.ToList();
            //Searching records from list using LINQ query  
            var TypeList = (from T in ObjList
                            where RemoveAcentuation(T.tipo).IndexOf(Prefix, StringComparison.OrdinalIgnoreCase) != -1
                            select new { T.tipo });
            var json = JsonConvert.SerializeObject(TypeList);
            return new TextResult(json, Request);
        }

        [System.Web.Http.HttpGet]
        public IHttpActionResult BrandsAutocomplete(string Prefix)
        {
            var ObjList = (from t in db.t_ecf_vehicles
                           group t by new { t.Marca }
                            into grp
                           select new
                           {
                               grp.Key.Marca
                           }).ToList();

            var BrandList = (from b in ObjList
                             where b.Marca != null && RemoveAcentuation(b.Marca).IndexOf(Prefix, StringComparison.OrdinalIgnoreCase) != -1
                             select new { b.Marca });
            var json = JsonConvert.SerializeObject(BrandList);
            return new TextResult(json, Request);
        }

        [System.Web.Http.HttpGet]
        public IHttpActionResult ModelsAutocomplete(string Prefix)
        {
            var ObjList = (from t in db.t_ecf_vehicles
                           group t by new { t.Modelo }
                            into grp
                           select new
                           {
                               grp.Key.Modelo
                           }).ToList();

            var ModelList = (from b in ObjList
                             where b.Modelo != null && RemoveAcentuation(b.Modelo).IndexOf(Prefix, StringComparison.OrdinalIgnoreCase) != -1
                             select new { b.Modelo });
            var json = JsonConvert.SerializeObject(ModelList);
            return new TextResult(json, Request);
        }

        [System.Web.Http.HttpGet]
        public IHttpActionResult CategoryAutocomplete(string Prefix)
        {
            var ObjList = (from t in db.t_ecf_vehicles
                           group t by new { t.Categoria }
                            into grp
                           select new
                           {
                               grp.Key.Categoria
                           }).ToList();

            var CategoryList = (from b in ObjList
                                where b.Categoria != null && RemoveAcentuation(b.Categoria).IndexOf(Prefix, StringComparison.OrdinalIgnoreCase) != -1
                                select new { b.Categoria });
            var json = JsonConvert.SerializeObject(CategoryList);
            return new TextResult(json, Request);
        }
        // GET: api/veiculos/?CarRegistration=""&year=""
        public IHttpActionResult GetList(string carRegistration, int year)
        {
            var total = db.t_ecf_lines_fuel
                           .Where(l => l.ExpenseDate.Value.Year == year && l.CarRegistration == carRegistration)
                           .GroupBy(m => new { m.ExpenseDate.Value.Year, m.ExpenseDate.Value.Month })
                           .Select(m =>
                           new
                           {
                               year = m.Key.Year,
                               month = m.Key.Month,
                               Amount = m.Sum(s => s.ExpenseAmount),
                               Mileage = ((float)((int)(m.Average(s => s.Avg) * 100)) / 100)
                           }).ToList();

            var json = JsonConvert.SerializeObject(total);
            return new TextResult(json, Request);
        }

        // GET: api/veiculos/5
        public IHttpActionResult Get(int id)
        {
            var result = (from s in db.t_ecf_header where s.id == id select s.RequestDate.ToString()).Single();
            return new TextResult(result, Request);
        }

        // POST: api/veiculos
        public void Post([FromBody]string value)
        {
        }

        // PUT: api/veiculos/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/veiculos/5
        public void Delete(int id)
        {
        }
    }
}

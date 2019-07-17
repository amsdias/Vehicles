using ClosedXML.Excel;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using Veiculos.Models;

namespace Veiculos.Controllers
{
    public class VehiclesController : Controller
    {
        private vdlwebcontrolEntities db = new vdlwebcontrolEntities();

        [HttpPost]
        public FileResult Export(int? empresa, int? departamento, FormCollection formcollection)
        {
            DataTable dt = new DataTable("Veículos");
            dt.Columns.AddRange(new DataColumn[21] { new DataColumn("Nº"),
                                            new DataColumn("Matrícula"),
                                            new DataColumn("Tipo"),
                                            new DataColumn("Marca"),
                                            new DataColumn("Modelo"),
                                            new DataColumn("Categoria"),
                                            new DataColumn("Combustível"),
                                            new DataColumn("Ano"),
                                            new DataColumn("Responsável"),
                                            new DataColumn("Utilizadores 9-18h"),
                                            new DataColumn("Utilizadores 18-9h"),
                                            new DataColumn("Nº Cartão Pagamento"),
                                            new DataColumn("Ref. Cartão Pagamento"),
                                            new DataColumn("Nº Cartão Frota"),
                                            new DataColumn("Ref. Cartão Frota"),                                            
                                            new DataColumn("Kms"),
                                            new DataColumn("Plafond Mensal"),
                                            new DataColumn("Obs."),
                                            new DataColumn("Empresa"),
                                            new DataColumn("Departamento"),
                                            new DataColumn("Activo")
            });

            var t_ecf_vehicles = db.t_ecf_vehicles.Include(t => t.t__companies).Include(t => t.t__departments);

            if (empresa > 0)
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Empresa == empresa);
            }

            if (departamento > 0)
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Departamento == departamento);
            }

            var listOfSelectedTypes = formcollection["SelectedTypes"];
            var tipos = db.t_ecf_vehicle_types.ToList();

            if (!String.IsNullOrEmpty(listOfSelectedTypes))
            {
                var arrayOfSelectedTypes = listOfSelectedTypes.Split(',');
                t_ecf_vehicles = t_ecf_vehicles.Where(item => arrayOfSelectedTypes.Any(stringToCheck => item.Tipo.Contains(stringToCheck)));
            }

            foreach (var vehicle in t_ecf_vehicles)
            {
                dt.Rows.Add(vehicle.No_, 
                    vehicle.Matricula,
                    vehicle.Tipo,
                    vehicle.Marca,
                    vehicle.Modelo,
                    vehicle.Categoria,
                    vehicle.Combustivel,
                    vehicle.Ano,
                    vehicle.Responsavel,
                    vehicle.Utilizadores_9_18,
                    vehicle.Utilizadores_18_9,
                    vehicle.NumCartaoPagamento,
                    vehicle.RefCartaoPagamento,
                    vehicle.NumCartaoFrota,
                    vehicle.RefCartaoFrota,                    
                    vehicle.Kilometros,
                    vehicle.PlafondMensal,
                    vehicle.Observacoes,
                    vehicle.t__companies.CompDescription,
                    vehicle.t__departments.department_pt,
                    vehicle.Activo
                    );
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "veiculos.xlsx");
                }
            }
        }

        // GET: vehicles
        public ActionResult Index(string searchString, int? empresa, int? departamento, string tipo, FormCollection formcollection, bool? filteractive, string sortOrder)
        {
            //companies to SHOW
            List<int> emprs = new List<int> { 1, 2, 4, 6 };

            //departments to HIDE
            List<int> deps = new List<int> { 11, 19, 20, 21, 22, 23, 24, 26, 27, 28, 29 };

            var listOfSelectedTypes = formcollection["SelectedTypes"];
            
            var empresas = db.t__companies.Where(e => emprs.Any(e2 => e2 == e.id)).Select(c => new Empresa { EmpresaId = c.id, EmpresaNome = c.CompDescription }).ToList();

            var departamentos = db.t__departments.Where(d => !deps.Any(p2 => p2 == d.id)).Select(d => new Departamento { DepartamentoId = d.id, DepartamentoNome = d.department_pt }).ToList();

            var tipos = db.t_ecf_vehicle_types.ToList();

            IEnumerable<string> types = db.t_ecf_vehicle_types.Select(d => d.tipo).ToList();

            var t_ecf_vehicles = db.t_ecf_vehicles.Include(t => t.t__companies).Include(t => t.t__departments);

            var finalList = new List<t_ecf_vehicles>();
            if (!String.IsNullOrEmpty(searchString))
            {
                var strings = searchString.Trim().Split(' ');
                t_ecf_vehicles = t_ecf_vehicles.Where(item => strings.Any(stringToCheck => item.Tipo.Contains(stringToCheck)));
            }
            if (!String.IsNullOrEmpty(listOfSelectedTypes))
            {
                var arrayOfSelectedTypes = listOfSelectedTypes.Split(',');
                t_ecf_vehicles = t_ecf_vehicles.Where(item => arrayOfSelectedTypes.Any(stringToCheck => item.Tipo.Contains(stringToCheck)));
            }
            
            if (empresa > 0)
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Empresa == empresa);
            }

            if (departamento > 0)
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Departamento == departamento);
            }

            if (!String.IsNullOrEmpty(tipo))
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Tipo.Contains(tipo));
            }

            string[] emptyArray = new string[] { };
            
            if (filteractive == true)
            {
                t_ecf_vehicles = t_ecf_vehicles.Where(v => v.Activo);
            }
           
            ViewBag.MatriculaSortParm = sortOrder == "matricula" ? "tipo_desc" : "matricula";
            ViewBag.TipoSortParm = sortOrder == "tipo" ? "tipo_desc" : "tipo";
            ViewBag.MarcaSortParm = sortOrder == "marca" ? "marca_desc" : "marca";
            ViewBag.AnoSortParm = sortOrder == "ano" ? "ano_desc" : "ano";
            ViewBag.PlafondSortParm = sortOrder == "plafond" ? "plafond_desc" : "plafond";

            switch (sortOrder)
            {
                case "matricula_desc":
                    t_ecf_vehicles = t_ecf_vehicles.OrderByDescending(s => s.Matricula);
                    break;
                case "matricula":
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.Matricula);
                    break;
                case "tipo":
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.Tipo);
                    break;
                case "tipo_desc":
                    t_ecf_vehicles = t_ecf_vehicles.OrderByDescending(s => s.Tipo);
                    break;
                case "marca":
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.Marca);
                    break;
                case "marca_desc":
                    t_ecf_vehicles = t_ecf_vehicles.OrderByDescending(s => s.Marca);
                    break;
                case "ano":
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.Ano);
                    break;
                case "ano_desc":
                    t_ecf_vehicles = t_ecf_vehicles.OrderByDescending(s => s.Ano);
                    break;
                case "plafond":
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.PlafondMensal);
                    break;
                case "plafond_desc":
                    t_ecf_vehicles = t_ecf_vehicles.OrderByDescending(s => s.PlafondMensal);
                    break;
                default:
                    t_ecf_vehicles = t_ecf_vehicles.OrderBy(s => s.ID);
                    break;
            }

            finalList = t_ecf_vehicles.ToList();

            if (filteractive == null) filteractive = false;
            var viewModel = new IndexViewModel { EmpresasList = empresas, VeiculosList = finalList, DepartamentosList = departamentos, TypesList = tipos, Types = types, SelectedTypes = !String.IsNullOrEmpty(listOfSelectedTypes) ? listOfSelectedTypes.Split(',') : emptyArray, isFilterActive = (bool)filteractive };

            return View(viewModel);
        }

        // GET: vehicles/Details/5
        public ActionResult Details(int? id, int? page)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            t_ecf_vehicles vehicle = t_ecf_vehicles.FindVehicle(id);
            if (vehicle == null)
            {
                return HttpNotFound();
            }

            var lines = db.t_ecf_lines_fuel.Where(l => l.CarRegistration == vehicle.Matricula).ToList();
            foreach (t_ecf_lines_fuel line in lines)
            {
                line.TimeStamp = db.t_ecf_header.FirstOrDefault(h => h.id == line.EcfFk).RequestDate;
            }

            int pageSize = 20;
            int pageNumber = (page ?? 1);
            var viewModel = new DetailsViewModel { Vehicle = vehicle, Lines = lines.ToPagedList(pageNumber, pageSize) };
            return View(viewModel);
        }

        // GET: vehicles/Create 
        public ActionResult Create()
        {
            ViewBag.Empresa = new SelectList(db.t__companies, "id", "CompDescription");
            ViewBag.Departamento = new SelectList(db.t__departments, "id", "department_pt");
            return View();
        }

        // POST: vehicles/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,No_,Matricula,Tipo,Marca,Modelo,Categoria,Combustivel,Ano,Empresa,Departamento,Responsavel,Utilizadores_9_18,Utilizadores_18_9,NumCartaoPagamento,RefCartaoPagamento,NumCartaoFrota,RefCartaoFrota,Kilometros,PlafondMensal,Observacoes,Car,Activo")] t_ecf_vehicles t_ecf_vehicles)
        {
            if (ModelState.IsValid)
            {
                t_ecf_vehicles.Car = t_ecf_vehicles.Matricula + " - " + t_ecf_vehicles.Marca + " " + t_ecf_vehicles.Modelo;
                db.t_ecf_vehicles.Add(t_ecf_vehicles);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.Empresa = new SelectList(db.t__companies, "id", "CompDescription", t_ecf_vehicles.Empresa);
            ViewBag.Departamento = new SelectList(db.t__departments, "id", "department_pt", t_ecf_vehicles.Departamento);
            return View(t_ecf_vehicles);
        }

        // GET: vehicles/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            t_ecf_vehicles vehicle = t_ecf_vehicles.FindVehicle(id);
            if (vehicle == null)
            {
                return HttpNotFound();
            }

            ViewBag.Empresa = new SelectList(db.t__companies, "id", "CompDescription", vehicle.Empresa);
            ViewBag.Departamento = new SelectList(db.t__departments, "id", "department_pt", vehicle.Departamento);
            return View(vehicle);
        }

        // POST: vehicles/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,No_,Matricula,Tipo,Marca,Modelo,Categoria,Combustivel,Ano,Empresa,Departamento,Responsavel,Utilizadores_9_18,Utilizadores_18_9,NumCartaoPagamento,RefCartaoPagamento,NumCartaoFrota,RefCartaoFrota,Kilometros,PlafondMensal,Observacoes,Car,Activo")] t_ecf_vehicles t_ecf_vehicles)
        {
            if (ModelState.IsValid)
            {
                t_ecf_vehicles.Car = t_ecf_vehicles.Matricula + " - " + t_ecf_vehicles.Marca + " " + t_ecf_vehicles.Modelo;
                db.Entry(t_ecf_vehicles).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.Empresa = new SelectList(db.t__companies, "id", "CompDescription", t_ecf_vehicles.Empresa);
            ViewBag.Departamento = new SelectList(db.t__departments, "id", "department_pt", t_ecf_vehicles.Departamento);
            return View(t_ecf_vehicles);
        }

        // GET: vehicles/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            t_ecf_vehicles t_ecf_vehicles = db.t_ecf_vehicles.Find(id);
            if (t_ecf_vehicles == null)
            {
                return HttpNotFound();
            }

            return View(t_ecf_vehicles);
        }

        // POST: vehicles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            t_ecf_vehicles t_ecf_vehicles = db.t_ecf_vehicles.Find(id);
            db.t_ecf_vehicles.Remove(t_ecf_vehicles);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

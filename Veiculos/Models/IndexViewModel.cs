using System.Collections.Generic;

namespace Veiculos.Models
{
    public class IndexViewModel
    {
        public List<Empresa> EmpresasList { get; set; }
        public List<Departamento> DepartamentosList { get; set; }
        public List<t_ecf_vehicles> VeiculosList { get; set; }
        public List<t_ecf_vehicle_types> TypesList { get; set; }
        public IEnumerable<string> Types { get; set; }
        public string[] SelectedTypes { get; set; }
        public bool isFilterActive { get; set; }
    }
}
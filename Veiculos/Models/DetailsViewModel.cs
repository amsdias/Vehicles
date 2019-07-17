using PagedList;

namespace Veiculos.Models
{
    public class DetailsViewModel
    {
        public t_ecf_vehicles Vehicle { get; set;}
        public IPagedList<t_ecf_lines_fuel> Lines { get; set; }
    }
}
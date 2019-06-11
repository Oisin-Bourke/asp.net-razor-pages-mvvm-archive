/*
The model class has the same bame as the razor page with .cs

The runtime looks for Razor Pages files in the Pages folder by default.

*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using RazorPagesHolland.Models;

using Microsoft.AspNetCore.Mvc.Rendering;

namespace RazorPagesHolland.Pages.Hollands
{
    public class AdminModel : PageModel
    {
        private readonly RazorPagesHollandContext _context;

        public AdminModel(RazorPagesHollandContext context)
        {
            _context = context;
        }

        public IList<Holland> Holland { get;set; }

        [BindProperty(SupportsGet = true)]
        public string SearchString { get; set; }//SearchString tied to Location
     
        public SelectList VesselNames { get; set; }
        [BindProperty(SupportsGet = true)]
        public string VesselNameSelect { get; set; }

        public SelectList DiveNames { get; set; }
        [BindProperty(SupportsGet = true)]
        public string DiveNameSelect { get; set; }

        public static string markers;//a class varaible to hold the markers


        /* When a request is made for the page this method initialize the state for the page.==> same as onGet while onPost handles form submissions*/
        public async Task OnGetAsync()
        {
            //LINQ query returns all holland video instances:
            var hollands = from h in _context.Holland
                           select h;

            //LINQ query gets a list of all vessel names from db context:
            IQueryable<string> vesselQuery = from h in _context.Holland
                                             orderby h.VesselName
                                             select h.VesselName;

            IQueryable<string> diveQuery =   from h in _context.Holland
                                             orderby h.ROVDiveName
                                             select h.ROVDiveName;

            if (!string.IsNullOrEmpty(SearchString))
            {
                hollands = hollands.Where(s => s.Location.Contains(SearchString));//LINQ query operator 'Contains'
            }

            if (!string.IsNullOrEmpty(VesselNameSelect))
            {
                hollands = hollands.Where(x => x.VesselName == VesselNameSelect);
            }

            if (!string.IsNullOrEmpty(DiveNameSelect))
            {
                hollands = hollands.Where(y => y.ROVDiveName == DiveNameSelect);
            }

            Holland = await hollands.ToListAsync();

            VesselNames = new SelectList(await vesselQuery.Distinct().ToListAsync());//project the distinct vessel names

            DiveNames = new SelectList(await diveQuery.Distinct().ToListAsync());

        }
    }
} 

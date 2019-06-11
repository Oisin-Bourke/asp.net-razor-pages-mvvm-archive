using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using RazorPagesHolland.Models;

namespace RazorPagesHolland.Pages.Hollands
{
    public class CreateModel : PageModel
    {
        private readonly RazorPagesHolland.Models.RazorPagesHollandContext _context;

        public CreateModel(RazorPagesHolland.Models.RazorPagesHollandContext context)
        {
            _context = context;
        }

        //The OnGet method initializes any state needed for the page
        public IActionResult OnGet()
        {
            return Page();
        }

        [BindProperty]//The Movie property uses the [BindProperty] attribute to opt-in to model binding. 
        public Holland Holland { get; set; }

        //The OnPostAsync method is run when the page posts form data
        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page(); //he Page method creates a PageResult object that renders the Create.cshtml page.
            }

            _context.Holland.Add(Holland);//add video data

            await _context.SaveChangesAsync();//save changes

            return RedirectToPage("./Admin");//after form submission return to index
        }
    }
}
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using RazorPagesHolland.Models;

namespace RazorPagesHolland.Pages.Hollands
{
    public class EditModel : PageModel
    {
        private readonly RazorPagesHolland.Models.RazorPagesHollandContext _context;

        public EditModel(RazorPagesHolland.Models.RazorPagesHollandContext context)
        {
            _context = context;
        }

        [BindProperty]
        public Holland Holland { get; set; }

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            Holland = await _context.Holland.FirstOrDefaultAsync(m => m.ID == id);

            if (Holland == null)
            {
                return NotFound();
            }

            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            _context.Attach(Holland).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!HollandExists(Holland.ID))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return RedirectToPage("./Admin");
        }

        private bool HollandExists(int id)
        {
            return _context.Holland.Any(e => e.ID == id);
        }

    }
}

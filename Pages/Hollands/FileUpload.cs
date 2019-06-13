using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using RazorPagesHolland.Models;

namespace RazorPagesHolland.Pages.Hollands
{
    public class FileUploadModel : PageModel
    {
        //porvides info about the webhosting environment:
        private IHostingEnvironment _environment;

        //db connection/session
        private readonly Models.RazorPagesHollandContext _context;

        public FileUploadModel(IHostingEnvironment environment, Models.RazorPagesHollandContext context)
        {
            _environment = environment;
            _context = context;
        }

        [BindProperty]
        public IFormFile Upload { get; set; }

        public Holland Holland { get; set; }

        public IList<Survey> Survey { get; private set; }

        static int theID;//the id passed to every survey instance (FK)

        //on page load do this with holland parameter id:
        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            else
            {
                theID = (int)id;
                Debug.WriteLine("THE ID IS:" + theID);
            }


            //make holland instance by id accessible and include item's surveys  
            Holland = await _context.Holland
                                    .Include(s => s.Surveys)
                                    .AsNoTracking()
                                    .FirstOrDefaultAsync(m => m.ID == id);


            if (Holland == null)
            {
                return NotFound();
            }

            //Survey = await _context.Survey.AsNoTracking().to

            return Page();
        }

        //action post method: 
        public async Task <IActionResult> OnPostAsync()
        {

            XSSFWorkbook xssfFile = null;//the excel file 

            try
            {
                //get the uploaded file string path from environment
                var filePath = Path.Combine(_environment.ContentRootPath, "uploads", Upload.FileName);

                //create/upload the file
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    await Upload.CopyToAsync(fileStream);
                    Debug.WriteLine("uploaded file");
                }

                //read the file
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    xssfFile = new XSSFWorkbook(file);
                    Debug.WriteLine("created xssf file");
                }

                //get the excel sheet:
                ISheet sheet = xssfFile.GetSheetAt(1);

                //iterate every 360th row (per hour) started from row 1 ***
                for (int row = 1; row <= sheet.LastRowNum; row += 360)
                {
                    DateTime dateTimeStamp = DateTime.MinValue;

                   // DateTime date = DateTime.MinValue;
                   // TimeSpan time = TimeSpan.MinValue;

                    int vesselId = 0;
                    float latitude = 0;
                    float longitude = 0;
                    int hollandId = theID;

                    Debug.WriteLine(dateTimeStamp + "*" + vesselId + "*" + latitude + "*" + longitude + "*" + hollandId);

                    //adapt code to actual survery tracks...
                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                    {
                        //new code for different data....get date and time from cells and add to date time
                        DateTime date = sheet.GetRow(row).GetCell(2).DateCellValue;
                        DateTime tempDate = sheet.GetRow(row).GetCell(3).DateCellValue;
                        TimeSpan time = tempDate.TimeOfDay;
                        dateTimeStamp = date + time;
                        
                        //dateTimeStamp = sheet.GetRow(row).GetCell(4).DateCellValue;
                        vesselId = (int)sheet.GetRow(row).GetCell(0).NumericCellValue;
                        latitude = (float)sheet.GetRow(row).GetCell(7).NumericCellValue;
                        longitude = (float)sheet.GetRow(row).GetCell(6).NumericCellValue;

                        Debug.WriteLine(dateTimeStamp + " | " + vesselId + " | " + latitude + " | " + longitude);
                    }
       
                    var survey = new Survey()
                    {
                        TimeInterval = dateTimeStamp,
                        VesselNumber = vesselId,
                        Latitude = latitude,
                        Longitude = longitude,
                        HollandID = hollandId
                    };

                    Debug.WriteLine("added this object" + survey);

                    //add each survey to the db and immediately save/persist to maintain order
                    _context.Survey.Add(survey);
                    await _context.SaveChangesAsync();
                }

            }
            catch (FileNotFoundException e)
            {
                Debug.WriteLine(e.Message);
            }
            catch (DbUpdateConcurrencyException e)
            {
                Debug.WriteLine(e.Message);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }

            /*
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
            */

            //return Page();

            return RedirectToPage("./Admin");
        }
    }
}
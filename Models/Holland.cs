using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


//defining the schema or 'data model' for Holland ROV Video data

namespace RazorPagesHolland.Models
{
    public class Holland
    {

        public int ID { get; set; }//primary key

        [Required]
        [Display(Name ="Vessel Name")]
        public string VesselName { get; set; }

        [Required]
        [Display(Name = "ROV Dive Name")]
        public string ROVDiveName { get; set; }

        [Required]
        public string Location { get; set; }

        //auto-required 
        [Display(Name = "Dive Date")]
        [DataType(DataType.Date)]
        public DateTime DiveDate { get; set; }

        [Required]
        [Display(Name = "Video")]
        public string VideoUrl { get; set; }

        //auto-required
        [Column(TypeName = "float(10, 6)")]
        public float Latitude { get; set; }

        //auto-required
        [Column(TypeName = "float(10, 6)")]
        public float Longitude { get; set; }

        //each video has many survey instances 
        public ICollection<Survey> Surveys { get; set; }

    }
}

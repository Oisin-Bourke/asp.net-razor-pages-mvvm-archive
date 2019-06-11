using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace RazorPagesHolland.Models
{
    public class Survey
    {
        public int SurveyID { get; set; }//PK

        public DateTime TimeInterval { get; set; }

        public int VesselNumber { get; set; }

        [Column(TypeName = "float(10, 6)")]
        public float Latitude { get; set; }

        [Column(TypeName = "float(10, 6)")]//only get 5 digits after decimal
        public float Longitude { get; set; }

        public int HollandID { get; set; }//FK

        public Holland Holland { get; set; }//Each survey instance has one holland video instance  


        public override string ToString()
        {
            return string.Format("[Survey: SurveyID={0}, TimeInterval={1}, VesselNumber={2}, Latitude={3}, Longitude={4}, HollandID={5}]", SurveyID, TimeInterval, VesselNumber, Latitude, Longitude, HollandID);
        }
    }
}

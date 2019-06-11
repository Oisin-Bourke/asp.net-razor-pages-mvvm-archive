using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Linq;


namespace RazorPagesHolland.Models
{
    public static class SeedData
    {
        public static void Initialize(IServiceProvider serviceProvider)
        {
            using (var context = new RazorPagesHollandContext(
            serviceProvider.GetRequiredService<
                DbContextOptions<RazorPagesHollandContext>>()))
            {
                //Look for movies
                if (context.Holland.Any())
                {
                    Console.WriteLine("db alreadys seeded");
                    return;//db alreadys seeded 

                }

                context.Holland.AddRange
                (

                    new Holland
                    {
                        VesselName = "Seed Data",
                        ROVDiveName = "Exploration",
                        Location = "North Atlantic Ocean",
                        DiveDate = DateTime.Parse("2018-01-01"),
                        VideoUrl = "aAcQPBcCsyY",
                        Latitude = 49.443101f,
                        Longitude = -10.770815f
                    },
                    new Holland
                    {
                        VesselName = "Seed Data",
                        ROVDiveName = "Exploration",
                        Location = "North Atlantic Ocean",
                        DiveDate = DateTime.Parse("2018-01-01"),
                        VideoUrl = "aAcQPBcCsyY",
                        Latitude = 49.443101f,
                        Longitude = -10.770815f
                    },
                    new Holland
                    {
                        VesselName = "Seed Data",
                        ROVDiveName = "Exploration",
                        Location = "North Atlantic Ocean",
                        DiveDate = DateTime.Parse("2018-01-01"),
                        VideoUrl = "aAcQPBcCsyY",
                        Latitude = 49.443101f,
                        Longitude = -10.770815f
                    },
                    new Holland
                    {
                        VesselName = "Seed Data",
                        ROVDiveName = "Exploration",
                        Location = "North Atlantic Ocean",
                        DiveDate = DateTime.Parse("2018-01-01"),
                        VideoUrl = "aAcQPBcCsyY",
                        Latitude = 49.443101f,
                        Longitude = -10.770815f
                    }

                    );
                context.SaveChanges();

            }
        }
    }
}

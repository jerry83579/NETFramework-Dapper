using Microsoft.SqlServer.Types;
using System;
using System.Collections.Generic;
using System.Data.Spatial;
using System.Linq;
using System.Spatial;
using System.Web;

namespace FileTest.Models
{
    public class Info
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Price { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public double? Lat { get; set; }
        public double? Longitude { get; set; }
        public SqlGeometry Location { get; set; }
        public string Category { get; set; }
    }
}
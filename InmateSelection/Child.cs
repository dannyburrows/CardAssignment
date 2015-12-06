using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InmateSelection
{
    public sealed class Child
    {
        public string Name { get; set; }
        public string DOC { get; set; }
        public string Facility { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public int SelectedCount { get; set; }

        public Child()
        {

        }

        public Child(string Name, string DOC, string Facility, string Address1, string Address2, string City, string State, string Zip)
        {
            this.Name = Name;
            this.DOC = DOC;
            this.Facility = Facility;
            this.Address1 = Address1;
            this.Address2 = Address2;
            this.City = City;
            this.State = State;
            this.Zip = Zip;
        }
    }
}

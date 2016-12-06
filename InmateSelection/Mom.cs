using System;
using System.Collections.Generic;
using System.Data;

namespace CardAssignment
{
    public sealed class Mom
    {
        public string Name { get; set; }
        public int CardsRequested { get; set; }
        public Child Child { get; set; }
        public List<Child> ChildrenToSendCards { get; set; }
        public bool HasParticipatingChild => Child != null && !Child.SkipChild;
        public int CardsNeededForChild => HasParticipatingChild ? Child.CardsNeeded : 0;

        public Mom()
        {
            ChildrenToSendCards = new List<Child>();
        }

        public Mom(DataRow Row)
        {
            ChildrenToSendCards = new List<Child>();

            if (Row["Cards"] == DBNull.Value || Row["Cards"].ToString().Contains("?"))
            {
                CardsRequested = 0;
            } else {
                CardsRequested = Convert.ToInt32(Row["Cards"]);
            }

            Name = Row["Mom"].ToString();
            if (!(Row["Child"] == DBNull.Value || string.IsNullOrEmpty(Row["Child"].ToString()) || Row["Child"].ToString().ToLower().Contains("no son")))
            {
                Child = new Child
                {
                    Name = Row["Child"].ToString(),
                    DOC = Row["DOC #"].ToString(),
                    Facility = Row["Facility"].ToString(),
                    Address1 = Row["Address #1"].ToString(),
                    Address2 = Row["Address #2"].ToString(),
                    City = Row["City"].ToString(),
                    State = Row["State"].ToString(),
                    Zip = Row["Zip"].ToString(),
                    SkipChild = (Row["SkipChild"].ToString().ToLower() == "x")
                };
            }
        }
    }
}

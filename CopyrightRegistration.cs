using System;

namespace CopyrightRegistration
{

    public class CopyrightRegistration
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime RegistrationDate { get; set; }
        public string Type { get; set; }

        public CopyrightRegistration(int id, string title, string author, DateTime registrationDate, string type)
        {
            Id = id;
            Title = title;
            Author = author;
            RegistrationDate = registrationDate;
            Type = type;
        }
    }
}

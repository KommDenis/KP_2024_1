using System;
using System.Collections.Generic;
using System.Linq;

namespace CopyrightRegistration
{
    // Клас для управління об'єктами авторських прав
    public class CopyrightRegistrationManager
    {
        private List<CopyrightRegistration> registrations;

        public CopyrightRegistrationManager()
        {
            registrations = new List<CopyrightRegistration>();
        }

        public void AddRegistration(CopyrightRegistration registration)
        {
            registrations.Add(registration);
        }

        // Методи повинні мати доступність, відповідну поверненому типу
        public List<CopyrightRegistration> GetAllRegistrations()
        {
            return registrations;
        }

        public void UpdateRegistration(int id, string title, string author, DateTime registrationDate, string type)
        {
            var existingRegistration = registrations.FirstOrDefault(r => r.Id == id);
            if (existingRegistration != null)
            {
                existingRegistration.Title = title;
                existingRegistration.Author = author;
                existingRegistration.RegistrationDate = registrationDate;
                existingRegistration.Type = type;
            }
        }

        public void RemoveRegistration(int id)
        {
            var registrationToRemove = registrations.FirstOrDefault(r => r.Id == id);
            if (registrationToRemove != null)
                registrations.Remove(registrationToRemove);
        }

        // Сортування за назвою
        public void SortByTitle()
        {
            registrations = registrations.OrderBy(r => r.Title).ToList();
        }

        // Пошук за назвою
        public List<CopyrightRegistration> SearchByTitle(string title)
        {
            return registrations.Where(r => r.Title.Contains(title)).ToList();
        }

        // Фільтрація за типом
        public List<CopyrightRegistration> FilterByType(string type)
        {
            return registrations.Where(r => r.Type.Equals(type, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        // Сортування за датою реєстрації
        public void SortByRegistrationDate()
        {
            registrations = registrations.OrderBy(r => r.RegistrationDate).ToList();
        }

        // Пошук за автором
        public List<CopyrightRegistration> SearchByAuthor(string author)
        {
            return registrations.Where(r => r.Author.Contains(author)).ToList();
        }
    }
}

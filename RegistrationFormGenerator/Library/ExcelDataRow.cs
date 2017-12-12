namespace RegistrationFormGenerator.Library
{
    class ExcelDataRow
    {
        public string Serial { get; set; }

        public string NameBengali { get; set; }
        public string NameEnglish { get; set; }
        public string DateOfBirth { get; set; }

        public string RegistrationNo { get; set; }
        public string RollNo { get; set; }

        public string SessionBengali { get; set; }
        public string SessionEnglish { get; set; }

        public string FatherNameBengali { get; set; }
        public string FatherNameEnglish { get; set; }

        public string MotherNameBengali { get; set; }
        public string MotherNameEnglish { get; set; }

        public string MobileNo { get; set; }

        public string PresentAddress { get; set; }
        public string PermanentAddress { get; set; }

        public int TemplateAutoSelectCode { get; set; }

        public string FacultyBengali { get; set; }
        public string FacultyEnglish { get; set; }

        public string DepertmentBengali { get; set; }
        public string DepertmentEnglish { get; set; }

        public string DegreeNameBengali { get; set; }
        public string DegreeNameEnglish { get; set; }

        public string AdmissionCancelled { get; set; }
        public string Comment { get; set; }
    }
}

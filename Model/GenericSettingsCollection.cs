using System.Collections;

namespace InstructorBriefcaseExtractor.Model
{
    public  class GenericSettingsCollection : CollectionBase
    {
        public GenericSettingsCollection()
        { }

        public GenericSettings Add(int Position,        string Description)
        {
            GenericSettings mvar = new GenericSettings
            {
                Position = Position,
                Description = Description
            };

            this.List.Add(mvar);
            return mvar;
        }
    }
}

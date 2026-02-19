using QED.GUS;

namespace VibrationVIEW_GUS
{
    public interface IGusExtended : IGus
    {
        string GUS_GetTestProfiles(string Filter);
    }
}

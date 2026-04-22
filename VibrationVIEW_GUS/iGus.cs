namespace VibrationVIEW_GUS
{
    public interface iGus
    {
        string GUS_Open_App(string app);
        string GUS_Scan_Devices();
        string GUS_GetDeviceInfo();
        string GUS_OpenDevice(string device);
        string GUS_CloseDevice(string device);
        void GUS_CloseApp();
        string GUS_PrepareTest(string testName);
        string GUS_StartTest();
        string GUS_StopTest();
        string GUS_PauseTest();
        string GUS_ContinueTest();
        string GUS_CloseTest();
        string GUS_GetStatus();
        string GUS_GetError();
        string GUS_GetInfo();
    }
}

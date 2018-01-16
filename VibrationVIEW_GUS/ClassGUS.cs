// VibrationVIEW_GUS
// GUS interface to VibrationVIEW controller
// Copyright (C) 2016  Vibration Research Corporation
//
// This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
//
// Project Dependencies:
//   VibrationVIEW (32 bit) - COM only supported (32 bit), COM+ NOT SUPPORTED (64 bit)


using System;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;
using System.Xml.Schema;
using System.Diagnostics;
using System.Text.RegularExpressions;
using QED.GUS;
using VibrationVIEWLib;

namespace VibrationVIEW_GUS
{
    public class GUS : IGus
    {
        enum TestTypes
        {
            TEST_SYSCHECK = 0,
            TEST_SINE = 1,
            TEST_RANDOM = 2,
            TEST_SOR = 3,
            TEST_SHOCK = 4,
            TEST_TRANSIENT = 5,
            TEST_REPLAY = 6,
            TEST_WAVEFORM = 7,
            TEST_CALIBRATION = 8,
            MAX_TEST_TYPES = 0
        };

        // COM interface to VibrationVIEW (COM only works x86 (32 bit)
        private VibrationVIEW _VibrationVIEWControl = null;

        // success returns GusConstants.CallReturnSucces
        // interface does not define fail, documentation indicates "ERR"
        const string CallReturnFAIL = "ERR";

        // BOOL false as defined by VibrationVIEW
        const int FALSE = 0;

        // VibrationVIEW Error codes
        int STOP_WAITING_FOR_BOX = 0x103A;

        // interface state variables
        private string _VibrationVIEWTestName = "";
        private string _VibrationVIEWDevice = "";
        private string _State = GusStatus.DeviceClosed;

        /// <summary>
        /// Release any create objects when the class is closed
        /// </summary>
        /// <returns></returns>
        ~GUS()
        {
            ReleaseVibrationVIEW();
        }
        /// <summary>
        /// The communication is closed. 
        /// Operating and communication software of the equipment is terminated.
        /// </summary>
        /// <returns>"ERR"/GusConstants.CallReturnSuccess</returns>
        public void GUS_CloseApp()
        {
            ReleaseVibrationVIEW();
        }

        /// <summary>
        /// The connection with the respective equipment is closed. 
        /// The equipment becomes available for other control applications. 
        /// The running process of the equipment is not influenced.
        /// </summary>
        /// <param name="device">As identified bu GUS_OpenDevice</param>
        /// <returns>"ERR"/GusConstants.CallReturnSuccess</returns>
        public string GUS_CloseDevice(string device)
        {
            if (_State == GusStatus.DeviceOpen)
            {
                // don't close mismatched software device
                if ((device == _VibrationVIEWDevice) || (device == ""))
                {
                    _State = GusStatus.DeviceClosed;
                    _VibrationVIEWDevice = "";
                    // we only have one device, don't really care if its open.
                    return GusConstants.CallReturnSuccess;
                }
            }

            // called from invalid state or invalid device
            return CallReturnFAIL;
        }

        /// <summary>
        /// The loaded test is closed. The equipment returns to the default state.
        /// The equipment is subsequently ready to receive a new PREPARE command.
        /// </summary>
        /// <returns>"ERR"/GusConstants.CallReturnSuccess</returns>
        public string GUS_CloseTest()
        {
            // defined in VibrationVIEW Resource.h file
            const int ID_TEST_CLOSE = 33050;
            UpdateRunningState();
            if ((_State == GusStatus.Ready) || (_State == GusStatus.Finished) || (_State == GusStatus.Error))
            {
                if (_VibrationVIEWTestName != "")
                {
                    try
                    {
                        // switch to the selected test
                        _VibrationVIEWControl.OpenTest(_VibrationVIEWTestName);
                        // and close the test
                        _VibrationVIEWControl.MenuCommand(ID_TEST_CLOSE);
                    }
                    catch (Exception)
                    {
                        return CallReturnFAIL;
                    }
                }

                _VibrationVIEWControl.set_TestType(vvTestType.VV_TEST_SYSCHECK);
                _VibrationVIEWTestName = "";
                _State = GusStatus.DeviceOpen;

                return GusConstants.CallReturnSuccess;
            }
            // invalid state
            return CallReturnFAIL;
        }

        /// <summary>
        /// The equipment goes in RUN mode again. The running test is continued.
        /// The run schedule, elapsed time, etc. of the interrupted test are not reset
        /// </summary>
        /// <returns>"ERR"/GusConstants.CallReturnSuccess</returns>
        public string GUS_ContinueTest()
        {
            UpdateRunningState();
            if (_State == GusStatus.Pause)
            {
                try
                {
                    // we should be able to resume the test
                    if (_VibrationVIEWControl.Running == FALSE)
                    {
                        _VibrationVIEWControl.ResumeTest();
                    }
                    else if (_VibrationVIEWControl.HoldLevel != FALSE)
                    {
                        const int ID_TEST_RUNTEST = 32870;
                        _VibrationVIEWControl.MenuCommand(ID_TEST_RUNTEST);
                    }
                    else
                    {
                        const int ID_TEST_ADVANCETONEXTLEVEL = 32896;
                        _VibrationVIEWControl.MenuCommand(ID_TEST_ADVANCETONEXTLEVEL);
                    }

                    // check if the test restarted
                    if (_VibrationVIEWControl.Running == FALSE)
                    {
                        // this should be caught in VibrationVIEW
                        _State = GusStatus.Error;
                        return CallReturnFAIL;
                    }
                    else
                    {
                        _State = GusStatus.Running;
                        return GusConstants.CallReturnSuccess;
                    }
                }
                catch (Exception)
                {
                    _State = GusStatus.Error;
                    return CallReturnFAIL;
                }
            }

            // called from invalid state or failed to start
            return CallReturnFAIL;
        }

        /// <summary>
        /// Read Equipment Properties
        /// “Get Device Info” will show all info from the device as available
        /// </summary>
        /// <returns>XML device info</returns>
        public string GUS_GetDeviceInfo()
        {
            // Must open device prior to running test
            try
            {
                String result;
                /*----------------------------------------------------------------------------------------------------------*/
                Encoding utf8 = new UTF8Encoding(false);                   // Set Streamwriter to UTF-8 

                XmlWriterSettings settings = new XmlWriterSettings();

                settings.Encoding = utf8;
                settings.Indent = true;
                settings.CheckCharacters = false;

                /*----------------------------------------------------------------------------------------------------------*/
                using (System.IO.MemoryStream xStreamData = new System.IO.MemoryStream())
                {
                    System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(xStreamData, settings);

                    //            xStreamData.Flush();

                    writer.WriteStartDocument(true);
                    writer.WriteStartElement("Device", "http://www.gus-interface.com/GusDeviceInfo");
                    writer.WriteAttributeString("xmlns", "xsi", "http://www.w3.org/2000/xmlns/", XmlSchema.InstanceNamespace);
                    writer.WriteAttributeString("xsi", "schemaLocation", null, "http://www.gus-interface.com/GusDeviceInfo GusDeviceInfo.xsd");

                    writer.WriteStartElement("Group");
                    writer.WriteAttributeString("Name", "DeviceInfo");

                    // Device Info               
                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Name");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "DeviceType");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Manufacturer");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "DeviceModel");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Address");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Remark");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteEndElement();

                    writer.WriteStartElement("Group");
                    writer.WriteAttributeString("Name", "ControlledValues");

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Control");
                    writer.WriteElementString("IsReadOnly", "true");

                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "Decimal");
                    string controlunit = _VibrationVIEWControl.get_ReportField("Control%f %s");
                    writer.WriteElementString("EngineeringUnit", controlunit.Substring(controlunit.IndexOf(" ") + 1));
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Demand");
                    writer.WriteElementString("IsReadOnly", "true");

                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "Decimal");

                    string demandunit = _VibrationVIEWControl.get_ReportField("Demand%f %s");
                    writer.WriteElementString("EngineeringUnit", demandunit.Substring(demandunit.IndexOf(" ") + 1));
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteEndElement();

                    writer.WriteStartElement("Group");
                    writer.WriteAttributeString("Name", "Testing");

                    writer.WriteStartElement("Attribute");
                    writer.WriteAttributeString("Name", "Stopcode");
                    writer.WriteElementString("IsReadOnly", "true");
                    writer.WriteStartElement("Type");
                    writer.WriteAttributeString("xsi", "type", null, "String");
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    int testtype = _VibrationVIEWControl.get_TestType();
                    if ((int)TestTypes.TEST_SHOCK == testtype)
                    {
                        writer.WriteStartElement("Attribute");
                        writer.WriteAttributeString("Name", "PulsesRun");
                        writer.WriteElementString("IsReadOnly", "true");
                        writer.WriteStartElement("Type");
                        writer.WriteAttributeString("xsi", "type", null, "Integer");
                        writer.WriteEndElement();
                        writer.WriteEndElement();

                        writer.WriteStartElement("Attribute");
                        writer.WriteAttributeString("Name", "PulsesScheduled");
                        writer.WriteElementString("IsReadOnly", "true");
                        writer.WriteStartElement("Type");
                        writer.WriteAttributeString("xsi", "type", null, "Integer");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    else
                    {
                        writer.WriteStartElement("Attribute");
                        writer.WriteAttributeString("Name", "TimeElapsedInTolerance");
                        writer.WriteElementString("IsReadOnly", "true");
                        writer.WriteStartElement("Type");
                        writer.WriteAttributeString("xsi", "type", null, "Integer");
                        writer.WriteElementString("EngineeringUnit", "Sec");
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteStartElement("Group");
                    writer.WriteAttributeString("Name", "Measurements");
                    for (int channelindex = 0; channelindex < _VibrationVIEWControl.HardwareInputChannels; channelindex++)
                    {
                        writer.WriteStartElement("Attribute");
                        writer.WriteAttributeString("Name", string.Format("Measurement{0}", channelindex + 1));
                        writer.WriteElementString("IsReadOnly", "true");
                        writer.WriteStartElement("Type");
                        writer.WriteAttributeString("xsi", "type", null, "Decimal");

                        string channel = _VibrationVIEWControl.get_ReportField(string.Format("Ch{0}%f %s", channelindex + 1));
                        writer.WriteElementString("EngineeringUnit", demandunit.Substring(channel.IndexOf(" ") + 1));
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();

                    /*----------------------------------------------------------------------------------------------------------*/
                    // Device End           
                    writer.WriteEndElement();

                    // Document End
                    writer.WriteEndDocument();
                    xStreamData.Flush();

                    writer.Close();

                    result = Encoding.Default.GetString(xStreamData.ToArray());

                    /*----------------------------------------------------------------------------------------------------------*/
                    xStreamData.Dispose();
                    Debug.WriteLine(result);

                    return result;
                }
            }

            catch (Exception e)
            {
                return CallReturnFAIL;
            }
        }
        /*----------------------------------------------------------------------------------------------------------*/

        /// <summary>
        /// Query of the equipment failure condition.
        /// Not required in minimum spec
        /// </summary>
        /// <returns>xml Text string with summary of all device errors in clear text (only devices and system errors, no test cancellation info etc)
        /// </returns>
        public string GUS_GetError()
        {
            try
            {
                // only return error if the fault is an abort
                if (_VibrationVIEWControl.Aborted != FALSE)
                {
                    int iReturn;
                    string csReturn;
                    _VibrationVIEWControl.Status(out csReturn, out iReturn);

                    return csReturn;
                }
                return GusConstants.CallReturnSuccess;
            }
            catch (Exception)
            {
                return CallReturnFAIL;
            }
        }

        /// <summary>
        /// Query of additional equipment status information.
        /// </summary>
        /// <returns>Arbitrary status string</returns>
        public string GUS_GetInfo()
        {
            try
            {
                String result;
                /*----------------------------------------------------------------------------------------------------------*/
                Encoding utf8 = new UTF8Encoding(false);                   // Set Streamwriter to UTF-8 

                XmlWriterSettings settings = new XmlWriterSettings();

                settings.Encoding = utf8;
                settings.Indent = true;
                settings.CheckCharacters = false;

                /*----------------------------------------------------------------------------------------------------------*/
                using (System.IO.MemoryStream xStreamData = new System.IO.MemoryStream())
                {
                    System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(xStreamData, settings);

                    //            xStreamData.Flush();

                    writer.WriteStartDocument(false);

                    writer.WriteStartElement("Device");

                    // Device Info               
                    writer.WriteStartElement("DeviceInfo");
                    writer.WriteElementString("Name", "VibrationVIEW_Default");
                    writer.WriteElementString("Manufacturer", "Vibration Research");
                    writer.WriteElementString("DeviceModel", "VR9500");
                    writer.WriteElementString("Address", VibrationViewSerialNumber());
                    writer.WriteElementString("Remark", "Test Interface");
                    writer.WriteEndElement();

                    writer.WriteStartElement("ControlledValues");
                    writer.WriteElementString("Control", _VibrationVIEWControl.get_ReportField("Control%.2f"));
                    writer.WriteElementString("Demand", _VibrationVIEWControl.get_ReportField("Demand%.2f"));

                    writer.WriteEndElement();

                    writer.WriteStartElement("Testing");
                    writer.WriteElementString("Stopcode", _VibrationVIEWControl.get_ReportField("Stopcode"));
                    int testtype = _VibrationVIEWControl.get_TestType();
                    if ((int)TestTypes.TEST_SHOCK == testtype)
                    {
                        writer.WriteElementString("PulsesRun", Regex.Match(_VibrationVIEWControl.get_ReportField("Pulses"), @"\d+").Value);
                        writer.WriteElementString("PulsesScheduled", Regex.Match(_VibrationVIEWControl.get_ReportField("Pulses"), @"(\d+)(?!.*\d)").Value);
                    }
                    else
                    {
                        string LevelTime = _VibrationVIEWControl.get_ReportField("LevelTime");
                        TimeSpan ts = TimeSpan.Parse(LevelTime);
                        writer.WriteElementString("TimeElapsedInTolerance", string.Format("{0}", ts.TotalSeconds));
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("Measurements");
                    for (int channelindex = 0; channelindex < _VibrationVIEWControl.HardwareInputChannels; channelindex++)
                    {
                        writer.WriteElementString(string.Format("Measurement{0}", channelindex + 1), _VibrationVIEWControl.get_ReportField(string.Format("Ch{0}%.2f", channelindex + 1)));
                    }
                    writer.WriteEndElement();

                    writer.WriteEndElement();

                    // Document End
                    writer.WriteEndDocument();
                    xStreamData.Flush();

                    writer.Close();

                    result = Encoding.Default.GetString(xStreamData.ToArray());

                    /*----------------------------------------------------------------------------------------------------------*/
                    xStreamData.Dispose();
                    Debug.WriteLine(result);

                    return result;
                }
            }
            catch (Exception e)
            {
                return CallReturnFAIL;
            }
        }

        /// <summary>
        /// Query of the equipment status: Ready, Stop, Pause, Run.
        /// </summary>
        /// <returns>GusStatus</returns>
        public string GUS_GetStatus()
        {
            UpdateRunningState();
            return _State;
            // omit State 6 : “Busy” In the standard, the state “Busy” 6 was introduced for any event where the device was “busy” doing
            //  something, not immediately related to an actual state. E.g. opening a file, writing data to disk, is
            //  busy with a transition,… This state has not been defined as a unique state, and as it has been
            //  interpreted, does not relate to a unique condition of a device. Hence, the busy state will be ignored
            //  by the higher level control program. It is not recommended to implement this state in the device,
            //  until this state has been defined uniquely. Then the state machine will be updated accordingly.
        }

        /// <summary>
        /// The hardware to communicate with is defined. 
        /// The operating and communication software is loaded. 
        /// The communication is initialized.
        /// </summary>
        /// <param name="device"></param>
        /// <returns></returns>
        public string GUS_OpenDevice(string device)
        {

            if ((_State == GusStatus.DeviceClosed) && (_VibrationVIEWControl != null))
            {
                // allow vibrationVIEW to fully initialize
                DateTime dtStartStatus = DateTime.Now;
                int iReturn;
                string csReturn;
                {
                    _VibrationVIEWControl.Status(out csReturn, out iReturn);
                    // the happy path is anything but STOP_WAITING_FOR_BOX
                    if (iReturn != STOP_WAITING_FOR_BOX)
                    {
                        // if we don't define the device use whatever box is connected
                        if (device == "")
                        {
                            device = VibrationViewSerialNumber();
                        }
                        else if (device != VibrationViewSerialNumber())
                        {
                            return CallReturnFAIL;
                        }
                        _State = GusStatus.DeviceOpen;
                        _VibrationVIEWDevice = device;
                        return GusConstants.CallReturnSuccess;
                    }
                }
                while ((DateTime.Now - dtStartStatus).Seconds < 10) ;

            }
            // called with invalid state or could not connect to a box
            return CallReturnFAIL;
        }

        /// <summary>
        /// The command “GUS_Open_App” will load the driver. The parameter of the command is the name of\
        /// the driver, as registered by the Windows® operating system.
        /// The device(s) that communicates by means of the selected driver is in the state “Closed” 9.
        /// </summary>
        /// <param name="app">Default VibrationVIEW - not checked</param>
        /// <returns>The response must be a string: “ACK: device serial/version number”</returns>
        public string GUS_Open_App(string app)
        {
            try
            {
                if (_VibrationVIEWControl == null)
                {
                    _VibrationVIEWControl = new VibrationVIEW();
                }
                if (_VibrationVIEWControl != null)
                {
                    _State = GusStatus.DeviceClosed;

                    return string.Format("{0}:{1}", GusConstants.CallReturnSuccess, _VibrationVIEWControl.SoftwareVersion.ToString());
                }
                else /* FAILED to create */
                {
                    return CallReturnFAIL;
                }
            }
            catch (Exception)
            {
                // software version failed to return?
                return CallReturnFAIL;
            }
        }

        /// <summary>
        /// The equipment goes in PAUSE mode. The running test goes on hold 
        /// but is not terminated. The elapsed time, 
        /// status of the runs schedule, etc. within the test remain unchanged.
        /// The equipment remains ready to continue the test or to stop the test.
        /// </summary>
        /// <returns></returns>
        public string GUS_PauseTest()
        {
            UpdateRunningState();
            if (_State == GusStatus.Running)
            {
                try
                {
                    _VibrationVIEWControl.StopTest();
                    if (_VibrationVIEWControl.CanResumeTest == FALSE)
                    {
                        // stop should normally allow resume
                        // if resume is NOT allowed, flag as error
                        _State = GusStatus.Error;
                        return CallReturnFAIL;
                    }

                    _State = GusStatus.Pause;
                    return GusConstants.CallReturnSuccess;
                }
                catch (Exception)
                {
                    return CallReturnFAIL;
                }
            }
            // called from invalid state
            return CallReturnFAIL;
        }

        /// <summary>
        /// A predefined test is selected and loaded. The initialization is executed (e.g. selfcheck)
        /// VV does not do self check like DACTRON.  We do have a "Pre-test" but its integrated into the run
        /// At the current time there is no way to advance from "Pre-test" to "Run" from ActiveX
        /// </summary>
        /// <param name="testName"></param>
        /// <returns></returns>
        public string GUS_PrepareTest(string testName)
        {
            // Must open device prior to running test
            if (_State == GusStatus.DeviceOpen)
            {
                _VibrationVIEWTestName = "";
                try
                {
                    _VibrationVIEWControl.OpenTest(testName);
                    _VibrationVIEWTestName = testName;
                    _State = GusStatus.Ready;
                    return GusConstants.CallReturnSuccess;
                }
                catch (Exception)
                {
                    _State = GusStatus.Error;
                    return CallReturnFAIL;
                }
            }
            // called from invalid state
            return CallReturnFAIL;
        }

        /// <summary>
        /// Search for available Equipment. The equipment identifier is returned.
        /// The equipment identifier return string is defined as xml
        /// not real clear what is supposed to be here, and the call is defined as <OPTIONAL>
        /// we shall chose to not implement
        /// </summary>
        /// <returns></returns>
        public string GUS_Scan_Devices()
        {
            return "";

        }

        /// <summary>
        /// The equipment goes in RUN mode. The loaded test is started and starts 
        /// to run until the end of test, or until the command STOP or PAUSE is 
        /// received, or until eventually a failure terminates the test.
        /// </summary>
        /// <returns></returns>
        public string GUS_StartTest()
        {
            UpdateRunningState();

            try
            {
                switch (_State)
                {
                    case GusStatus.Ready:
                        if (_VibrationVIEWTestName == "")
                        {
                            // out of sequence
                            // this could happen with operator intervention, 
                            //  or application starting after VV has test running
                            //  if GUS_PrepareTest was never run, we can not start the test
                            //  assume error state
                            _State = GusStatus.Error;
                            return CallReturnFAIL;
                        }

                        _VibrationVIEWControl.RunTest(_VibrationVIEWTestName);

                        // make sure we started!
                        if (_VibrationVIEWControl.Running == FALSE)
                        {
                            _State = GusStatus.Error;
                            return CallReturnFAIL;
                        }
                        _State = GusStatus.PreTestRunning;
                        return GusConstants.CallReturnSuccess;
                    default:
                        // called from invalid state
                        return CallReturnFAIL;
                }
            }
            catch (Exception)
            {
                _State = GusStatus.Error;
                return CallReturnFAIL;
            }
        }

        /// <summary>
        /// The equipment goes in READY mode. The running test goes on hold and is terminated. 
        /// All control parameters (elapsed time, run schedule, etc.) 
        /// are reset and the equipment is ready for a new START. 
        /// The status is identical as after the PREPARE command.
        /// </summary>
        /// <returns></returns>
        public string GUS_StopTest()
        {
            UpdateRunningState();
            try
            {
                switch (_State)
                {
                    case GusStatus.Pause:
                    case GusStatus.Running:
                    case GusStatus.PreTestRunning:
                        _VibrationVIEWControl.StopTest();
                        _State = GusStatus.Ready;
                        break;
                    case GusStatus.Error:
                    case GusStatus.Finished:
                        _State = GusStatus.Ready;
                        break;
                    default:
                        // called with invalid state
                        return CallReturnFAIL;
                }

                return GusConstants.CallReturnSuccess;
            }
            catch (Exception)
            {
                return CallReturnFAIL;
            }
        }

        /// <summary>
        /// Check if we stopped running.
        /// Anytime we look at our RUNNING state we need to see if we aborted or finished
        /// </summary>
        /// <returns></returns>
        private void UpdateRunningState()
        {
            if (_State == GusStatus.PreTestRunning)
            {
                try
                {
                    // fix up our state if was running, now not running because faulted, completed, or stopped from GUI
                    if (_VibrationVIEWControl.Running == FALSE)
                    {
                        if (_VibrationVIEWControl.Aborted == FALSE)
                        {
                            _State = GusStatus.Finished;
                        }
                        else
                        {
                            _State = GusStatus.Error;
                        }
                    }
                    else
                    {
                        if (_VibrationVIEWControl.Starting == FALSE)
                        {
                            _State = GusStatus.Running;
                        }

                    }
                }
                catch (Exception)
                {
                    _State = GusStatus.Error;
                }
            }

            if (_State == GusStatus.Running)
            {
                try
                {
                    if (IsSchedulePause())
                    {
                        _State = GusStatus.Pause;
                    }

                    // fix up our state if was running, now not running because faulted, completed, or stopped from GUI
                    if (_VibrationVIEWControl.Running == FALSE)
                    {
                        if (_VibrationVIEWControl.Aborted == FALSE)
                        {
                            _State = GusStatus.Finished;
                        }
                        else
                        {
                            _State = GusStatus.Error;
                        }
                    }
                }
                catch (Exception)
                {
                    _State = GusStatus.Error;
                }
            }
            else
            {
                // if not running state, fix up the state to running if the controller is running because crashed, started from GUI, remote started
                try
                {
                    if ((_VibrationVIEWControl != null) && (_VibrationVIEWDevice != "") && (_VibrationVIEWControl.Running != FALSE))
                    {
                        if (_VibrationVIEWControl.Starting == FALSE)
                        {
                            if (IsSchedulePause())
                            {
                                _State = GusStatus.Pause;
                            }
                            else
                            {
                                _State = GusStatus.Running;
                            }
                        }
                        else
                        {
                            if (IsSchedulePause())
                            {
                                _State = GusStatus.Pause;
                            }
                            else
                            {
                                _State = GusStatus.PreTestRunning;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    _State = GusStatus.Error;
                }
            }
        }

        private bool IsSchedulePause()
        {
            int iReturn;
            string csReturn;
            _VibrationVIEWControl.Status(out csReturn, out iReturn);
            if ((iReturn & 0xff) == 0x31) // WAIT_FOR_OPERATOR
            {
                return true;
            }
            if ((iReturn & 0xff) == 0x31) // WAIT_FOR_OPERATOR
            {
                return true;
            }
            if (FALSE != _VibrationVIEWControl.HoldLevel)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// format VibrationVIEW VR9500 serial number as shown on back of BOX
        /// </summary>
        /// <returns></returns>
        private string VibrationViewSerialNumber()
        {
            try
            {
                string serialnumber = string.Format("{0,8:X}", _VibrationVIEWControl.HardwareSerialNumber);
                return serialnumber;
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Release COM ref to VibrationVIEW
        /// clear out any state associated with the released VibrationVIEW
        /// </summary>
        /// <returns></returns>
        private void ReleaseVibrationVIEW()
        {
            // delete our interface into VibrationVIEW
            if (_VibrationVIEWControl != null)
            {
                // Should not have to do this, but VV not destroyed unless I explicitly release it.
                Marshal.ReleaseComObject(_VibrationVIEWControl);
                _VibrationVIEWControl = null;
            }

            // reset our state variables
            _VibrationVIEWTestName = "";
            _VibrationVIEWDevice = "";
            _State = GusStatus.DeviceClosed;
        }
    }
}

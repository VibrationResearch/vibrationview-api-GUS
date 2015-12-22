using System;
using System.Text;
using QED.GUS;
using VibrationVIEWLib;
using System.Runtime.InteropServices;

namespace VibrationVIEW_GUS
{
    public class GUS:IGus
    {
			VibrationVIEW VibrationVIEWControl=null;

			// success returns GusConstants.CallReturnSucces
			// interface does not define fail, documentation indicates "ERR"
			const string CallReturnFAIL = "ERR";
			
			// BOOL false as defined by VibrationVIEW
			const int FALSE = 0;

			// VibrationVIEW Error codes
			int STOP_WAITING_FOR_BOX = 0x103A;
			
			// interface state variables
			private string VibrationVIEWTestName = "";
			private string VibrationVIEWDevice = "";
			private string state = GusStatus.DeviceClosed;

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
				if (state == GusStatus.DeviceOpen)
				{
					// don't close mismatched software device
					if ((device == VibrationVIEWDevice) || (device == ""))
					{
						state = GusStatus.DeviceClosed;
						VibrationVIEWDevice = "";
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
				UpdateRunningState();
				if ((state == GusStatus.Ready) || (state == GusStatus.Finished))
				{
					// we don't explicitly unload tests.
					VibrationVIEWTestName = "";
					state = GusStatus.DeviceOpen;
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
				if (state == GusStatus.Pause)
				{
					try
					{
						// we should be able to resume the test
						VibrationVIEWControl.ResumeTest();

						// check if the test restarted
						if (VibrationVIEWControl.Running == FALSE)
						{
							return CallReturnFAIL;
							// this should be caught in VibrationVIEW
							state = GusStatus.Error;
						}
						else
						{
							state = GusStatus.Running;
							return GusConstants.CallReturnSuccess;
						}
					}
					catch (Exception)
					{
						state = GusStatus.Error;
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
				//not implemented .. Per Statusmachine_comments_Version08 (Dec -2015) The commands “GUS_GetDeviceInfo” and GUS_GetInfo” will need further discussion in the work
				//  group of the GUS.
				return "";
			}

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
					if (VibrationVIEWControl.Aborted != FALSE)
					{
						int iReturn;
						string csReturn;
						VibrationVIEWControl.Status(out csReturn, out iReturn);

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
				//not implemented .. Per Statusmachine_comments_Version08 (Dec -2015) The commands “GUS_GetDeviceInfo” and GUS_GetInfo” will need further discussion in the work
				//  group of the GUS.
				return "";
			}

			/// <summary>
			/// Query of the equipment status: Ready, Stop, Pause, Run.
			/// </summary>
			/// <returns>GusStatus</returns>
			public string GUS_GetStatus()
			{
				UpdateRunningState();
				return state;
				// omit State 2: "PreTestRunning" The state diagram defines it as a dactron pre-test, which we don't do
				//  activeX has no concept of pretest.  We could see if pre-test is enabled using a PARAM, 
				//  but the interface has no way to transition from pre-test mode to run mode.
				//
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
				
				if ((state == GusStatus.DeviceClosed) && (VibrationVIEWControl != null))
				{
					// allow vibrationVIEW to fully initialize
					DateTime dtStartStatus = DateTime.Now;
					int iReturn;
					string csReturn;
					{
						VibrationVIEWControl.Status(out csReturn, out iReturn);
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
							state = GusStatus.DeviceOpen;
							VibrationVIEWDevice = device;
							return GusConstants.CallReturnSuccess;
						}
					}
					while ( (DateTime.Now - dtStartStatus).Seconds < 10);

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
					if (VibrationVIEWControl == null)
					{
						VibrationVIEWControl = new VibrationVIEW();
					}
					if (VibrationVIEWControl != null)
					{
							state = GusStatus.DeviceClosed;
							return string.Format("{0}:{1}", GusConstants.CallReturnSuccess, VibrationVIEWControl.SoftwareVersion.ToString());
					}
					else /* FAILED to create */
					{
						return CallReturnFAIL;
					}
				}
				catch(Exception)
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
				if(state == GusStatus.Running)
				{
					try
					{
						VibrationVIEWControl.StopTest();
						if (VibrationVIEWControl.CanResumeTest == FALSE)
						{
							// stop should normally allow resume
							// if resume is NOT allowed, flag as error
							state = GusStatus.Error;
							return CallReturnFAIL;
						}

						state = GusStatus.Pause;
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
				if(state == GusStatus.DeviceOpen)
				{
					VibrationVIEWTestName = "";
					try
					{
						VibrationVIEWControl.OpenTest(testName);
						VibrationVIEWTestName = testName;
						state = GusStatus.Ready;
						return GusConstants.CallReturnSuccess;
					}
					catch(Exception)
					{
						state = GusStatus.ProjLoadFailed;
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

				if ((state == GusStatus.Finished) || (state == GusStatus.Ready))
				{
					try
					{
						if (state == GusStatus.Finished)
						{
							state = GusStatus.Ready;
						}
						else
						{
							VibrationVIEWControl.RunTest(VibrationVIEWTestName);
							if (VibrationVIEWControl.Running == FALSE)
							{
								state = GusStatus.Error;
								return CallReturnFAIL;
							}
							state = GusStatus.Running;
						}
						return GusConstants.CallReturnSuccess;
					}
					catch (Exception)			
					{
						state = GusStatus.Error;
						return CallReturnFAIL;
					}
				}
				else
				{
					// called from invalid state
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
					switch (state)
					{
						case GusStatus.Pause:
						case GusStatus.Ready:
						case GusStatus.Error:
						case GusStatus.Running:
							state = GusStatus.Finished;
							break;
						case GusStatus.ProjLoadFailed:
							state = GusStatus.DeviceOpen;
							break;
						default:
							// called with invalid state
							return CallReturnFAIL;
					}

					VibrationVIEWControl.StopTest();
					return GusConstants.CallReturnSuccess;
				}
				catch(Exception)
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
				if (state == GusStatus.Running)
				{
					try
					{
						if (VibrationVIEWControl.Running == FALSE)
						{
							if (VibrationVIEWControl.Aborted == FALSE)
							{
								state = GusStatus.Finished;
							}
							else
							{
								state = GusStatus.Error;
							}
						}
					}
					catch (Exception)
					{
						state = GusStatus.Error;
					}
				}
			}
			/// <summary>
			/// format VibrationVIEW VR9500 serial number as shown on back of BOX
			/// </summary>
			/// <returns></returns>
			private string VibrationViewSerialNumber()
			{
				try
				{
					string serialnumber = string.Format("{0,8:X}", VibrationVIEWControl.HardwareSerialNumber);
					return serialnumber;
				}
				catch(Exception)
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
				if (VibrationVIEWControl != null)
				{
					// Should not have to do this, but VV not destroyed unless I explicitly release it.
					Marshal.ReleaseComObject(VibrationVIEWControl);
					VibrationVIEWControl = null;
				}

				// reset our state variables
				VibrationVIEWTestName = "";
				VibrationVIEWDevice = "";
				state = GusStatus.DeviceClosed;
			}
		}
}

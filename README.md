# VibrationVIEW GUS Interface

A GUS (Generic Universal Standard) interface implementation for [VibrationVIEW](https://vibrationresearch.com/vibrationview/) vibration controllers by Vibration Research Corporation.

GUS is a standardized interface specification that enables higher-level control software to operate vibration test equipment from different manufacturers through a common API. This project implements the GUS interface for VibrationVIEW, allowing GUS-compliant test automation systems to control VibrationVIEW-connected shaker controllers (e.g., VR9700, VR10500).

## Architecture

The solution communicates with VibrationVIEW via its COM (ActiveX) interface.

```
GUS Client (e.g., GusTestInterface)
        |
    IGusExtended / IGus  (QED.GUS standard interface)
        |
    GUS class  (this project - VibrationVIEW_GUS.dll)
        |
    VibrationVIEW COM interface  (VibrationVIEWLib)
        |
    VibrationVIEW application / VR hardware
```

### Interfaces

- **`IGus`** (from `GUSInterface.dll`) - The standard GUS interface defining the core command set for vibration test equipment control.
- **`IGusExtended`** (this project) - Extends `IGus` with VibrationVIEW-specific additions such as `GUS_GetTestProfiles`.

## GUS State Machine

The GUS specification defines a state machine that governs equipment control flow:

```
Open_App --> DeviceClosed
                |
           OpenDevice
                |
           DeviceOpen
                |
           PrepareTest
                |
             Ready  <-- StopTest
            /     \
       StartTest   CloseTest --> DeviceOpen
          |
    PreTestRunning
          |
       Running  <---  ContinueTest
       /     \
  PauseTest   (test ends)
     |            |
   Pause      Finished / Error
```

## GUS Commands

| Command | Description |
|---|---|
| `GUS_Open_App` | Load the VibrationVIEW COM driver and initialize communication |
| `GUS_OpenDevice` | Connect to a specific VR controller by serial number |
| `GUS_CloseDevice` | Disconnect from the controller |
| `GUS_CloseApp` | Terminate communication and release COM objects |
| `GUS_PrepareTest` | Load a test profile by filename |
| `GUS_StartTest` | Start the loaded test |
| `GUS_StopTest` | Stop the running test and reset to Ready |
| `GUS_PauseTest` | Pause the running test (can be resumed) |
| `GUS_ContinueTest` | Resume a paused test |
| `GUS_CloseTest` | Close the loaded test and return to DeviceOpen |
| `GUS_GetStatus` | Query the current state (Ready, Running, Pause, Finished, Error, etc.) |
| `GUS_GetInfo` | Query device information and live measurements as XML |
| `GUS_GetDeviceInfo` | Query device property schema as XML |
| `GUS_GetError` | Query equipment error/abort information |
| `GUS_Scan_Devices` | Search for available equipment (optional, not implemented) |
| `GUS_GetTestProfiles` | List available test profiles from the profiles directory as XML |

### GUS_GetTestProfiles

Returns the contents of `c:\vibrationview\profiles` as XML. Accepts a filter parameter that can be either a file glob pattern or a test type name:

| Filter | Maps to |
|---|---|
| `sine` | `*.vsp` |
| `random` | `*.vrp` |
| `shock` | `*.vkp` |
| `datareplay` | `*.vfp` |
| `*.vsp` (etc.) | Used as-is |

The filter is case-insensitive. Example response:

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<TestProfiles>
  <Profile>
    <Name>my_sine_test.vsp</Name>
  </Profile>
  <Profile>
    <Name>another_test.vsp</Name>
  </Profile>
</TestProfiles>
```

## Return Values

- **`"ACK"`** - Success (`GusConstants.CallReturnSuccess`)
- **`"ERR"`** - Failure
- `GUS_Open_App` returns `"ACK:<version>"` on success

## Solution Structure

```
VibrationVIEW_GUS.sln
|
+-- VibrationVIEW_GUS/          Main library (VibrationVIEW_GUS.dll)
|   +-- ClassGUS.cs              GUS interface implementation
|   +-- IGusExtended.cs          Extended interface adding GUS_GetTestProfiles
|
+-- VibrationVIEW_GUS.Tests/    Unit tests (MSTest)
|   +-- GUS_GetTestProfilesTests.cs
|
+-- Gus Interface/              External GUS specification files and DLLs
|   +-- GUSInterface.dll         Standard IGus interface definition
|   +-- GusTestInterface.exe     GUS test client application
|
+-- SetupVibrationVIEW_Gus/     WiX installer project
```

## Requirements

- Windows operating system (compatible with Windows 10 and Windows 11)
- VibrationVIEW software installed
- VibrationVIEW automation option (VR9604) - OR - VibrationVIEW may be run in Simulation mode without any additional hardware or software
- Visual Studio 2022
- .NET Framework 4.5.2
- x86 (32-bit) build target (COM limitation)

## Building

Open `VibrationVIEW_GUS.sln` in Visual Studio and build. The project must be compiled as x86 since VibrationVIEW only exposes a 32-bit COM interface.

## License

This project is licensed under the MIT License with VibrationVIEW Attribution. See the [LICENSE](LICENSE) file for details.

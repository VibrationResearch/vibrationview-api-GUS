using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;

namespace VibrationVIEW_GUS
{
    /// <summary>
    /// Drop-in diagnostics to pin down why the cast to
    /// VibrationVIEWLib.VibrationVIEW fails with E_NOINTERFACE (0x80004002).
    ///
    /// Usage: inside GUS_Open_App, right BEFORE the failing cast, do:
    ///
    ///     object raw = Activator.CreateInstance(type);   // your existing line
    ///     GusComDiagnostics.Report(raw, type);           // <-- add this
    ///     var vv = (VibrationVIEWLib.VibrationVIEW)raw;   // the line that throws
    ///
    /// Report(...) never throws; it prints everything to the Output/Debug window
    /// (and Console) so you can read it even when the cast blows up right after.
    /// </summary>
    public static class GusComDiagnostics
    {
        // The strong interface IID your interop is asking for (from your error).
        private const string TARGET_IID = "7B945DAC-0BD4-47A8-BA6D-4EFB8C48EA8B";

        private static readonly System.Text.StringBuilder _logBuffer = new System.Text.StringBuilder();

        /// <summary>
        /// Returns and clears the accumulated diagnostic log.
        /// </summary>
        public static string FlushLog()
        {
            string result = _logBuffer.ToString();
            _logBuffer.Clear();
            return result;
        }

        // IID_IDispatch — every scriptable/automation object should expose this.
        private static readonly Guid IID_IDispatch =
            new Guid("00020400-0000-0000-C000-000000000046");

        public static void Report(object comObj, Type coclassType)
        {
            Log("===== GUS COM DIAGNOSTICS =====");

            if (comObj == null) { Log("comObj is NULL — CreateInstance returned nothing."); return; }
            Log("Runtime type       : " + comObj.GetType().FullName);
            Log("Is real COM object : " + Marshal.IsComObject(comObj));

            IntPtr pUnk = IntPtr.Zero;
            try
            {
                pUnk = Marshal.GetIUnknownForObject(comObj);

                // 1) Does the object expose IDispatch at all?
                bool hasDispatch = QueryInterface(pUnk, IID_IDispatch);
                Log("Supports IDispatch : " + hasDispatch);

                // 2) Does it expose the exact strong interface the interop wants?
                bool hasTarget = QueryInterface(pUnk, new Guid(TARGET_IID));
                Log("Supports {" + TARGET_IID + "} : " + hasTarget);

                // 3) If IDispatch works, ask the object what interface it REALLY is.
                if (hasDispatch)
                    DumpActualInterface(comObj);
            }
            catch (Exception ex)
            {
                Log("Diagnostic probe error: " + ex.Message);
            }
            finally
            {
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }

            // 4) What does the registry actually point the CLSID/IID at?
            DumpRegistry(coclassType);

            Log("===== INTERPRETATION =====");
            Log("IDispatch=YES but target IID=NO  -> interop/typelib VERSION MISMATCH.");
            Log("   Regenerate VibrationVIEWLib from the INSTALLED VibrationVIEW typelib.");
            Log("Both NO                          -> object is not the VV automation");
            Log("   coclass, or proxy/stub not registered (bitness/marshaling).");
            Log("Target IID missing from registry -> run VibrationVIEW.exe /regserver (admin).");
            Log("================================");
        }

        private static bool QueryInterface(IntPtr pUnk, Guid iid)
        {
            IntPtr pOut = IntPtr.Zero;
            try
            {
                int hr = Marshal.QueryInterface(pUnk, ref iid, out pOut);
                return hr == 0 && pOut != IntPtr.Zero;
            }
            finally
            {
                if (pOut != IntPtr.Zero) Marshal.Release(pOut);
            }
        }

        // Uses IDispatch::GetTypeInfo to read the object's real interface name + GUID.
        private static void DumpActualInterface(object comObj)
        {
            try
            {
                IDispatch disp = (IDispatch)comObj;
                int count;
                disp.GetTypeInfoCount(out count);
                Log("IDispatch type info count : " + count);
                if (count == 0) return;

                System.Runtime.InteropServices.ComTypes.ITypeInfo ti;
                disp.GetTypeInfo(0, 0, out ti);

                IntPtr pAttr;
                ti.GetTypeAttr(out pAttr);
                try
                {
                    var attr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)
                        Marshal.PtrToStructure(pAttr,
                            typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
                    Log("ACTUAL interface GUID     : {" + attr.guid + "}");

                    string name, doc, help;
                    int ctx;
                    ti.GetDocumentation(-1, out name, out doc, out ctx, out help);
                    Log("ACTUAL interface name     : " + name);
                }
                finally { ti.ReleaseTypeAttr(pAttr); }
            }
            catch (Exception ex)
            {
                Log("Could not read type info: " + ex.Message);
            }
        }

        private static void DumpRegistry(Type coclassType)
        {
            Log("Process bitness           : " + (IntPtr.Size == 8 ? "64-bit" : "32-bit"));

            try
            {
                // CLSID of the coclass we instantiated.
                object[] attrs = coclassType.GetCustomAttributes(typeof(GuidAttribute), false);
                if (attrs.Length > 0)
                {
                    string clsid = "{" + ((GuidAttribute)attrs[0]).Value + "}";
                    Log("CoClass CLSID             : " + clsid);
                    Log("LocalServer32 (32-bit)    : " +
                        ReadReg(RegistryView.Registry32, @"CLSID\" + clsid + @"\LocalServer32", null));
                    Log("LocalServer32 (64-bit)    : " +
                        ReadReg(RegistryView.Registry64, @"CLSID\" + clsid + @"\LocalServer32", null));
                }
            }
            catch (Exception ex) { Log("CLSID lookup error: " + ex.Message); }

            string iidKey = @"Interface\{" + TARGET_IID + "}";

            Log("--- 32-bit registry ---");
            Log("Interface entry           : " + (ReadReg(RegistryView.Registry32, iidKey, null) ?? "<<< MISSING >>>"));
            Log("  ProxyStubClsid32        : " + ReadReg(RegistryView.Registry32, iidKey + @"\ProxyStubClsid32", null));
            Log("  TypeLib                 : " + ReadReg(RegistryView.Registry32, iidKey + @"\TypeLib", null));

            Log("--- 64-bit registry ---");
            Log("Interface entry           : " + (ReadReg(RegistryView.Registry64, iidKey, null) ?? "<<< MISSING >>>"));
            Log("  ProxyStubClsid32        : " + ReadReg(RegistryView.Registry64, iidKey + @"\ProxyStubClsid32", null));
            Log("  TypeLib                 : " + ReadReg(RegistryView.Registry64, iidKey + @"\TypeLib", null));
        }

        private static string ReadReg(RegistryView view, string subkey, string valueName)
        {
            try
            {
                using (RegistryKey hkcr = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                using (RegistryKey k = hkcr.OpenSubKey(subkey))
                    return k?.GetValue(valueName ?? "")?.ToString() ?? "<not set>";
            }
            catch { return "<read error>"; }
        }

        private static string ReadHKCR(string subkey, string valueName)
        {
            try
            {
                using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(subkey))
                    return k?.GetValue(valueName ?? "")?.ToString() ?? "<not set>";
            }
            catch { return "<read error>"; }
        }

        private static void Log(string s)
        {
            _logBuffer.AppendLine(s);
            Debug.WriteLine(s);
            Console.WriteLine(s);
        }

        [ComImport, Guid("00020400-0000-0000-C000-000000000046"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatch
        {
            void GetTypeInfoCount(out int count);
            void GetTypeInfo(int index, int lcid,
                out System.Runtime.InteropServices.ComTypes.ITypeInfo typeInfo);
            // Remaining vtable slots (GetIDsOfNames, Invoke) intentionally omitted —
            // we only call the first two, which is enough to read type info.
        }
    }
}
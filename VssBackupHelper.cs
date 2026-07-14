using System;
using System.Runtime.InteropServices;

[ComImport, Guid("507C37B4-CF5B-4e95-B0AF-14EB9767467E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IVssAsync
{
    [PreserveSig] int Cancel();
    [PreserveSig] int Wait(int dwMilliseconds = -1);
    [PreserveSig] int QueryStatus(out int pHrResult, out int pReserved);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct VSS_SNAPSHOT_PROP
{
    public Guid m_SnapshotId;
    public Guid m_SnapshotSetId;
    public int m_lSnapshotsCount;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszSnapshotDeviceObject;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszOriginalVolumeName;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszOriginatingMachine;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszServiceMachine;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszExposedName;
    [MarshalAs(UnmanagedType.LPWStr)] public string m_pwszExposedPath;
    public Guid m_ProviderId;
    public int m_lSnapshotAttributes;
    public long m_tsCreationTimestamp;
    public int m_eStatus;
}

[ComImport, Guid("665c1d5f-c218-414d-a05d-7fef5f9d5c86"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IVssBackupComponents
{
    // 1-5
    void GetWriterComponentsCount(out uint pcComponents);
    void GetWriterComponents(uint index, out IntPtr ppWriter);
    void InitializeForBackup([MarshalAs(UnmanagedType.BStr)] string bstrXML);
    void SetBackupState(bool bSelectComponents, bool bBackupBootableSystemState, int backupType, bool bPartialFileSupport);
    void InitializeForRestore([MarshalAs(UnmanagedType.BStr)] string bstrXML);

    // 6-10
    void SetRestoreState(int restoreType);
    void GatherWriterMetadata(out IVssAsync ppAsync);
    void GetWriterMetadataCount(out uint pcWriters);
    void GetWriterMetadata(uint iWriter, out Guid pidInstance, out IntPtr ppMetadata);
    void FreeWriterMetadata();

    // 11-15
    void AddComponent(Guid instanceId, Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName);
    void PrepareForBackup(out IVssAsync ppAsync);
    void AbortBackup();
    void GatherWriterStatus(out IVssAsync ppAsync);
    void GetWriterStatusCount(out uint pcWriters);

    // 16-20
    void FreeWriterStatus();
    void GetWriterStatus(uint iWriter, out Guid pidInstance, out Guid pidWriter, [MarshalAs(UnmanagedType.BStr)] out string pbstrWriter, out int pws, out int phr);
    void SetBackupSucceeded(Guid instanceId, Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, bool bSucceeded);
    void SetBackupOptions(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszBackupOptions);
    void SetSelectedForRestore(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, bool bSelectedForRestore);

    // 21-25
    void SetRestoreOptions(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszRestoreOptions);
    void SetAdditionalRestores(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, bool bAdditionalRestores);
    void SetPreviousBackupStamp(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszPreviousBackupStamp);
    void SaveAsXML([MarshalAs(UnmanagedType.BStr)] out string pbstrXML);
    void BackupComplete(out IVssAsync ppAsync);

    // 26-30
    void AddAlternativeLocationMapping(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszPath, [MarshalAs(UnmanagedType.LPWStr)] string wszFilespec, bool bRecursive, [MarshalAs(UnmanagedType.LPWStr)] string wszDestination);
    void AddRestoreSubcomponent(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszSubcomponentLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszSubcomponentName, bool bRepair);
    void SetFileRestoreStatus(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, int status);
    void AddNewTarget(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [MarshalAs(UnmanagedType.LPWStr)] string wszPath, [MarshalAs(UnmanagedType.LPWStr)] string wszFileName, bool bRecursive, [MarshalAs(UnmanagedType.LPWStr)] string wszAlternatePath);
    void SetRangesFilePath(Guid writerId, int ct, [MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, uint iPartialFile, [MarshalAs(UnmanagedType.LPWStr)] string wszRangesFile);

    // 31-35
    void PreRestore(out IVssAsync ppAsync);
    void PostRestore(out IVssAsync ppAsync);
    void SetContext(int lContext);
    void StartSnapshotSet(out Guid pSnapshotSetId);
    void AddToSnapshotSet([MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, Guid ProviderId, out Guid pidSnapshot);

    // 36-40
    void DoSnapshotSet(out IVssAsync ppAsync);
    void DeleteSnapshots(Guid SourceObjectId, int SourceObjectType, bool bForceDelete, out int plDeletedSnapshots, out Guid pNondeletedSnapshotID);
    void ImportSnapshots(out IVssAsync ppAsync);
    void BreakSnapshotSet(Guid SnapshotSetId);
    void GetSnapshotProperties(Guid SnapshotId, IntPtr pProp);

    // 41-45
    void Query(Guid QueriedObjectId, int QueriedObjectType, int ReturnedObjectType, out IntPtr ppEnum);
    void IsVolumeSupported(Guid ProviderId, [MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, out bool pbSupportedByThisProvider);
    void DisableWriterClasses(ref Guid rgWriterClassId, uint cWriterClassId);
    void EnableWriterClasses(ref Guid rgWriterClassId, uint cWriterClassId);
    void DisableWriterInstances(ref Guid rgWriterInstanceId, uint cWriterInstanceId);
}

public class Program
{
    [DllImport("vssapi.dll", EntryPoint = "CreateVssBackupComponentsInternal", CallingConvention = CallingConvention.StdCall)]
    public static extern int CreateVssBackupComponents(out IVssBackupComponents ppBackup);

    [DllImport("vssapi.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern void VssFreeSnapshotProperties(IntPtr pProp);

    public static bool Succeeded(int hr)
    {
        return hr >= 0;
    }

    public static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("Usage: VssBackupHelper.exe <VolumePath> [Context]");
            return 1;
        }

        string volumePath = args[0];
        if (!volumePath.EndsWith("\\"))
        {
            volumePath += "\\";
        }

        int contextVal = 0; // Default context Backup (0)
        if (args.Length > 1)
        {
            string contextStr = args[1];
            if (contextStr.Equals("ClientAccessible", StringComparison.OrdinalIgnoreCase))
            {
                contextVal = 29;
            }
            else if (contextStr.Equals("ClientAccessibleWriters", StringComparison.OrdinalIgnoreCase))
            {
                contextVal = 13;
            }
            else if (contextStr.Equals("AppRollback", StringComparison.OrdinalIgnoreCase))
            {
                contextVal = 9;
            }
            else if (contextStr.Equals("Backup", StringComparison.OrdinalIgnoreCase))
            {
                contextVal = 0;
            }
        }

        IVssBackupComponents backupComponents = null;
        try
        {
            int hr = CreateVssBackupComponents(out backupComponents);
            if (!Succeeded(hr) || backupComponents == null)
            {
                Console.Error.WriteLine("Failed to create VssBackupComponents. HR: 0x" + hr.ToString("X"));
                return hr;
            }

            backupComponents.InitializeForBackup(null);
            backupComponents.SetContext(contextVal);

            backupComponents.SetBackupState(true, true, 1, false); // 1 = VSS_BT_FULL

            IVssAsync asyncObj = null;
            
            backupComponents.GatherWriterMetadata(out asyncObj);
            if (asyncObj != null)
            {
                asyncObj.Wait(-1);
                int queryHr = 0;
                int reserved = 0;
                asyncObj.QueryStatus(out queryHr, out reserved);
                Marshal.ReleaseComObject(asyncObj);
                if (!Succeeded(queryHr))
                {
                    Console.Error.WriteLine("GatherWriterMetadata failed. HR: 0x" + queryHr.ToString("X"));
                    return queryHr;
                }
            }

            Guid snapshotSetId;
            backupComponents.StartSnapshotSet(out snapshotSetId);

            Guid snapshotId;
            backupComponents.AddToSnapshotSet(volumePath, Guid.Empty, out snapshotId);

            backupComponents.PrepareForBackup(out asyncObj);
            if (asyncObj != null)
            {
                asyncObj.Wait(-1);
                int prepHr = 0;
                int reserved = 0;
                asyncObj.QueryStatus(out prepHr, out reserved);
                Marshal.ReleaseComObject(asyncObj);
                if (!Succeeded(prepHr))
                {
                    Console.Error.WriteLine("PrepareForBackup failed. HR: 0x" + prepHr.ToString("X"));
                    return prepHr;
                }
            }

            backupComponents.DoSnapshotSet(out asyncObj);
            if (asyncObj != null)
            {
                asyncObj.Wait(-1);
                int snapHr = 0;
                int reserved = 0;
                asyncObj.QueryStatus(out snapHr, out reserved);
                Marshal.ReleaseComObject(asyncObj);
                if (!Succeeded(snapHr))
                {
                    Console.Error.WriteLine("DoSnapshotSet failed. HR: 0x" + snapHr.ToString("X"));
                    return snapHr;
                }
            }

            int propSize = Marshal.SizeOf(typeof(VSS_SNAPSHOT_PROP));
            IntPtr pProp = Marshal.AllocHGlobal(propSize);
            try
            {
                backupComponents.GetSnapshotProperties(snapshotId, pProp);
                VSS_SNAPSHOT_PROP prop = (VSS_SNAPSHOT_PROP)Marshal.PtrToStructure(pProp, typeof(VSS_SNAPSHOT_PROP));
                
                Console.WriteLine("SUCCESS");
                Console.WriteLine("SnapshotID: {" + prop.m_SnapshotId.ToString() + "}");
                Console.WriteLine("DeviceObject: " + prop.m_pwszSnapshotDeviceObject);
                Console.WriteLine("SnapshotSetID: {" + prop.m_SnapshotSetId.ToString() + "}");
            }
            finally
            {
                VssFreeSnapshotProperties(pProp);
                Marshal.FreeHGlobal(pProp);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error: " + ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
            return -1;
        }
        finally
        {
            if (backupComponents != null)
            {
                Marshal.ReleaseComObject(backupComponents);
            }
        }

        return 0;
    }
}

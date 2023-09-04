using System.Runtime.InteropServices;
using FsrmLib;

namespace RootFolderPermissionsTest;
/// <summary>
/// for more info visit
/// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-fsrm/a6620ce9-e026-4a20-9ba2-e56280e230e0
/// </summary>
public class WindowsServerQuotaManager : IDisposable
{
    private readonly IFsrmQuotaManager _quotaManager = new FsrmQuotaManager();

    public IFsrmQuota? GetQuotaIfExists(string path)
    {
        try
        {
            return _quotaManager.GetQuota(path);
        }
        catch (COMException e)
        {
            var errorMessage = e.HResult switch
            {
                unchecked((int)0x80045301) => "The specified quota could not be found",
                unchecked((int)0x80045304) => "The quota for the specified path could not be found",
                unchecked((int)0x80045306) =>
                    "The content of the path parameter exceeds the maximum length of 4,000 characters",
                unchecked((int)0x80070057) =>
                    "The path parameter is NULL or the quota parameter is NULL",
                _ => "Unknown error"
            };

            if (e.HResult == unchecked((int)0x80045301))
                return null;

            throw new Exception(errorMessage);
        }
    }


    public void CreateOrUpdateQuota(string path, long limitInBytes)
    {
        try
        {
            var quota = GetQuotaIfExists(path) ?? _quotaManager.CreateQuota(path);
            quota.QuotaLimit = limitInBytes;
            quota.Description = "Created By BCLOUD system";
            quota.Commit();
        }
        catch (COMException e)
        {
            var errorMessage = e.HResult switch
            {
                unchecked((int)0x80045303) => "The quota for the specified path already exists",
                unchecked((int)0x80070057) => "One of the quota parameters is NULL",
                _ => "Unknown error"
            };

            throw new Exception(errorMessage);
        }
    }


    public void Dispose()
    {
        Marshal.ReleaseComObject(_quotaManager);
        Marshal.FinalReleaseComObject(_quotaManager);
    }
}
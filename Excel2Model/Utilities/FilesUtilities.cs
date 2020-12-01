using System.IO;

namespace Excel2Model.Utilities
{
    public static class FilesUtilities
    {
        public static bool FileExists(string filePath)
        {
            try
            {
                Path.GetFullPath(filePath);
            }
            catch
            {
                return false;
            }
            return true;
        }
    }
}

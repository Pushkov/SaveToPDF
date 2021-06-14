using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace SaveToPDF
{
    public class ROTHelper
    {
        #region APIs  

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport("ole32.dll")]
        private static extern int ProgIDFromCLSID([In()] ref Guid clsid, [MarshalAs(UnmanagedType.LPWStr)] out string lplpszProgID);

        #endregion

        #region Public Methods  

        /// <summary>  
        /// Converts a COM class ID into a prog id.  
        /// </summary>  
        /// <param name="progID">The prog id to convert to a class id.</param>  
        /// <returns>Returns the matching class id or the prog id if it wasn't found.</returns>  
        public static string ConvertProgIdToClassId(string progID)
        {
            Guid testGuid;
            try
            {
                CLSIDFromProgIDEx(progID, out testGuid);
            }
            catch
            {
                try
                {
                    CLSIDFromProgID(progID, out testGuid);
                }
                catch
                {
                    return progID;
                }
            }
            return testGuid.ToString().ToUpper();
        }

        /// <summary>  
        /// Converts a COM class ID into a prog id.  
        /// </summary>  
        /// <param name="classID">The class id to convert to a prog id.</param>  
        /// <returns>Returns the matching class id or null if it wasn't found.</returns>  
        public static string ConvertClassIdToProgId(string classID)
        {
            Guid testGuid = new Guid(classID.Replace("!", ""));
            string progId = null;
            try
            {
                ProgIDFromCLSID(ref testGuid, out progId);
            }
            catch (Exception)
            {
                return null;
            }
            return progId;
        }

        /// <summary>  
        /// Get a snapshot of the running object table (ROT).  
        /// </summary>  
        /// <param name="filter">The filter to apply to the list (nullable).</param>  
        /// <returns>A hashtable of the matching entries in the ROT</returns>  
        public static Dictionary<string, object> GetActiveObjectList(string filter)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();

            IntPtr numFetched = new IntPtr();
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            IBindCtx ctx;
            CreateBindCtx(0, out ctx);

            ctx.GetRunningObjectTable(out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                if (filter == null || filter.Length == 0 || runningObjectName.IndexOf(filter) != -1)
                {
                    object runningObjectVal;
                    runningObjectTable.GetObject(monikers[0], out runningObjectVal);
                    result[runningObjectName] = runningObjectVal;
                }
            }

            return result;
        }

        /// <summary>  
        /// Returns an object from the ROT, given a prog Id.  
        /// </summary>  
        /// <param name="progId">The prog id of the object to return.</param>  
        /// <returns>The requested object, or null if the object is not found.</returns>  
        public static object GetActiveObject(string progId)
        {
            // Convert the prog id into a class id  
            string classId = ConvertProgIdToClassId(progId);

            IRunningObjectTable runningObjectTable = null;
            IEnumMoniker monikerEnumerator = null;
            IBindCtx ctx = null;
            try
            {
                IntPtr numFetched = new IntPtr();
                // Open the running objects table.  
                CreateBindCtx(0, out ctx);
                ctx.GetRunningObjectTable(out runningObjectTable);
                runningObjectTable.EnumRunning(out monikerEnumerator);
                monikerEnumerator.Reset();
                IMoniker[] monikers = new IMoniker[1];

                // Iterate through the results  
                while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
                {

                    string runningObjectName;
                    monikers[0].GetDisplayName(ctx, null, out runningObjectName);
                    if (runningObjectName.IndexOf(classId) != -1)
                    {
                        // Return the matching object  
                        object objReturnObject;
                        runningObjectTable.GetObject(monikers[0], out objReturnObject);
                        return objReturnObject;
                    }
                }
                return null;
            }
            finally
            {
                // Free resources  
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);
                if (monikerEnumerator != null)
                    Marshal.ReleaseComObject(monikerEnumerator);
                if (ctx != null)
                    Marshal.ReleaseComObject(ctx);
            }
        }

        #endregion
    }
}
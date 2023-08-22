namespace VisualStudioFileOpenTool
{
    #region

    using System;
    using System.Collections;
    using System.Runtime.InteropServices;
    using EnvDTE;
    using System.Runtime.InteropServices.ComTypes;

   // warning CS0618: “UCOMIBindCtx”已过时:“Use System.Runtime.InteropServices.ComTypes.IBindCtx instead.http://go.microsoft.com/fwlink/?linkid=14202”
    #endregion
    public class ComHelper2
    {
        [DllImport("ole32.dll")]
        // public static extern int CreateBindCtx(int reserved, out UCOMIBindCtx ppbc);
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        /// <summary>
        /// Get a table of the currently running instances of the Visual Studio .NET IDE.
        /// </summary>
        /// <returns>A hashtable mapping the name of the IDE in the running object table to the corresponding DTE object</returns>
        public static Hashtable GetIDEInstances()
        {
            var runningIDEInstances = new Hashtable();
            var runningObjects = GetRunningObjectTable();

            var rotEnumerator = runningObjects.GetEnumerator();
            while (rotEnumerator.MoveNext())
            {
                var candidateName = (string)rotEnumerator.Key;
                if (!candidateName.StartsWith("!VisualStudio.DTE"))
                {
                    continue;
                }

                var ide = rotEnumerator.Value as _DTE;
                if (ide == null)
                {
                    continue;
                }

                runningIDEInstances[candidateName] = ide;
            }

            return runningIDEInstances;
        }

        public static ICollection GetRunningInstances()
        {
            return GetIDEInstances().Values;
        }

        // warning CS0618: “UCOMIRunningObjectTable”已过时:“Use System.Runtime.InteropServices.ComTypes.IRunningObjectTable instead. http://go.microsoft.com/fwlink/?linkid=14202”        
        [DllImport("ole32.dll")]
        // public static extern int GetRunningObjectTable(int reserved, out UCOMIRunningObjectTable prot);
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        /// <summary>
        /// Get a snapshot of the running object table (ROT).
        /// </summary>
        /// <returns>A hashtable mapping the name of the object in the ROT to the corresponding object</returns>
        [STAThread]
        public static Hashtable GetRunningObjectTable()
        {
            var result = new Hashtable();

       //     int numFetched;
            // UCOMIRunningObjectTable runningObjectTable;
            IRunningObjectTable runningObjectTable;
// warning CS0618: “UCOMIEnumMoniker”已过时:“Use System.Runtime.InteropServices.ComTypes.IEnumMoniker instead. http://go.microsoft.com/fwlink/?linkid=14202”            
            // UCOMIEnumMoniker monikerEnumerator;
            IEnumMoniker monikerEnumerator;
// warning CS0618: “UCOMIMoniker”已过时:“Use System.Runtime.InteropServices.ComTypes.IMoniker instead. http://go.microsoft.com/fwlink/?linkid=14202”            
            // var monikers = new UCOMIMoniker[1];
            var monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();
            
            
          //  IntPtr iptr = new IntPtr(numFetched);
            IntPtr fetched = IntPtr.Zero;
            // while (monikerEnumerator.Next(1, monikers, out numFetched) == 0)
            while (monikerEnumerator.Next(1, monikers, fetched) == 0)
            {
                // UCOMIBindCtx ctx;
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                result[runningObjectName] = runningObjectVal;
            }

            return result;
        }
    }
}

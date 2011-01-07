//-----------------------------------------------------------------------
// <copyright file="WmiContext.cs">
//     This code is published under the Microsoft Public License 
//     (Ms-PL). A copy of the license should be distributed with the code. 
//     It can also be found at the project website: 
//     http://www.CodePlex.com/linq2wmi. This notice, the author's name, 
//     and all copyright notices must remain intact in all applications, 
//     documentation, and source files.
// </copyright>
// 
// <author name="Joshua DeLong" date="12/14/2010 9:39:29 AM" />
//-----------------------------------------------------------------------

namespace LinqToWmi.Core.WMI
{
    #region Usings

    using System;
    using System.Collections;
    using System.IO;
    using System.Management;

    #endregion

    /// <summary>
    /// Represents an WMI context, and is responsible for opening, closing and
    /// querying to the WMI  We are now using the WMIContext both as a context
    /// and as a session (IE. Context.CreateSession()), this should be 2
    /// separate entities later on but for time sake, I'll just put them
    /// together.
    /// </summary>
    public class WmiContext : IDisposable
    {
        #region Usings

        private ManagementScope _managementScope;

        private WmiQueryBuilder builder;
        private string host;
        private TextWriter _log;

        #endregion

        #region Constructors/Deconstructors

        /// <summary>
        /// Initializes a new instance of the <see cref="WmiContext"/> class.
        /// </summary>
        public WmiContext() : this(@"\\.")
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WmiContext"/> class.
        /// </summary>
        /// <param name="host">The host.</param>
        public WmiContext(string host)
        {
            this.builder = new WmiQueryBuilder(this);
            this.host = host;
        }

        #endregion

        #region Properties

        public TextWriter Log
        {
            set { this._log = value; }
        }

        /// <summary>
        /// Defines the scope of the WMI query management object.
        /// </summary>
        public ManagementScope ManagementScope
        {
            get
            {
                if (this._managementScope == null)
                {
                    this.Connect();
                }

                return this._managementScope;
            }
        }

        #endregion

        /// <summary>
        /// Makes sure the connection exists and is created.
        /// </summary>
        private void EnsureConnectionCreated()
        {
            if (this._managementScope == null)
            {
                this.Connect();
            }
        }

        /// <summary>
        /// Creates in-fact a new WMI Query object.
        /// </summary>
        public WmiQuery<T> Source<T>()
        {
            return new WmiQuery<T>(this);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing,
        /// releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (this._managementScope != null)
            {
                this.Disconnect();
            }
        }

        /// <summary>
        /// Executes an WMI query (which is actually an wrapped expression).
        /// </summary>
        /// <param name="query">The WMI query.</param>
        /// <returns>A generic <see cref="IEnumerator"/> for the specified query
        /// context.</returns>
        internal IEnumerator ExecuteWmiQuery(IWmiQuery query)
        {
            this.EnsureConnectionCreated();
            var wmiQueryStatement = this.builder.BuildQuery(query);

            this.WriteLog(wmiQueryStatement);

            var objQuery = new ObjectQuery(wmiQueryStatement);
            var wmiSearcher = new ManagementObjectSearcher(this._managementScope, objQuery);

            // Generate a new generic type.
            var genericType = typeof(WmiObjectEnumerator<>).MakeGenericType(query.Type);

            var genericCollection = genericType.GetConstructor(new[] { typeof(ManagementObjectCollection) });

            return (IEnumerator)genericCollection.Invoke(new object[] { wmiSearcher.Get() });
        }

        /// <summary>
        /// Connect and create an new management scope.
        /// </summary>
        private void Connect()
        {
            this._managementScope = new ManagementScope(this.host, new ConnectionOptions());
        }

        /// <summary>
        /// Disconnect.
        /// </summary>
        private void Disconnect()
        {
        }

        /// <summary>
        /// Helper function to write text data to a log.
        /// </summary>
        /// <param name="txt">The text to be logged.</param>
        private void WriteLog(string txt)
        {
            if (this._log != null)
            {
                this._log.WriteLine(@"LOG: " + txt);
            }
        }
    }
}
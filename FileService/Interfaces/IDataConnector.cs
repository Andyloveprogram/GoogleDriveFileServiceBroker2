using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

using SourceCode.SmartObjects.Services.ServiceSDK.Attributes;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace SourceCode.Interfaces
{
    /// <summary>
    /// A ServiceBroker responsible for brokering communications between the K2 platform and an underlying system or technology.
    /// You do not need to make any changes to this class
    /// Note: this interface is specifically for Static Service Brokers that do not require the Execute() method
    /// </summary>
    interface IDataConnector : IDisposable
    {
        #region Methods

        #region void GetConfiguration()
        /// <summary>
        /// Gets the configuration from the service instance and stores the retrieved configuration in local variables for later use.
        /// </summary>
        void GetConfiguration();
        #endregion

        #region void SetupConfiguration()
        /// <summary>
        /// Sets up the required configuration parameters in the service instance. When a new service instance is registered for this ServiceBroker, the configuration parameters are surfaced to the appropriate tooling. The configuration parameters are provided by the person registering the service instance.
        /// </summary>
        void SetupConfiguration();
        #endregion

        #region void SetupService()
        /// <summary>
        /// Sets up the service instance's default name, display name, and description.
        /// </summary>
        void SetupService();
        #endregion

        #region void DescribeSchema()
        /// <summary>
        /// Describes the schema of the underlying data and services to the K2 platform.
        /// </summary>
        void DescribeSchema();
        #endregion

        #endregion
    }
}
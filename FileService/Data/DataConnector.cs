using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

using SourceCode.Interfaces;

namespace SourceCode.SmartObjects.Services
{
    /// <summary>
    /// An implementation of static IDataConnector responsible for interacting with an underlying system or technology. 
    /// The purpose of this class is to expose and represent the underlying data and services as Service Objects 
    /// which are, in turn, consumed by K2 SmartObjects
    /// TODO: implement the Interface Methods with your own code
    /// </summary>
    class DataConnector : IDataConnector
    {
        #region Class Level Fields

        #region Constants
        /// <summary>
        /// Constant for the Type Mappings configuration lookup in the service instance.
        /// </summary>
        private const string TYPEMAPPINGS = "Type Mappings";
        #endregion

        #region Private Fields
        /// <summary>
        /// Local serviceBroker variable.
        /// </summary>
        private ServiceAssemblyBase serviceBroker = null;

        /// <summary>
        /// Sample configuration values for the service instance
        /// defined by the SetupConfiguration() method and set by the GetConfiguration() method
        /// </summary>
        private const string credName = "OAuth 2.0 Client ID Path";
        private string credPath = string.Empty;

        //private string optionalConfigurationValue = string.Empty;

        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Instantiates a new DataConnector.
        /// </summary>
        /// <param name="serviceBroker">The ServiceBroker.</param>
        public DataConnector(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        #endregion

        #region Methods

        #region void SetupConfiguration()
        /// <summary>
        /// Sets up the required configuration parameters for the service instance. 
        /// When a new service instance is registered for this ServiceBroker, the configuration parameters are surfaced to the registration tool. 
        /// The configuration values are provided by the person registering the service instance. 
        /// You can mark the configuration properties as "required" and also provide a default value
        /// </summary>
        public void SetupConfiguration()
        {
            //TODO: Add the service instance configuration values here
            serviceBroker.Service.ServiceConfiguration.Add(credName, true, null);

        }
        #endregion

        #region void GetConfiguration()
        /// <summary>
        /// Retrieves the configuration from the service instance and stores the retrieved configuration in local variables for later use.
        /// </summary>
        public void GetConfiguration()
        {
            //TODO: Add code to retrieve the service instance configuration values
            credPath = serviceBroker.Service.ServiceConfiguration[credName].ToString();

        }
        #endregion

        #region void SetupService()
        /// <summary>
        /// Sets up the service instance's default name, display name, and description.
        /// </summary>
        public void SetupService()
        {
            //TODO: Add code to set up the service instance name, display name and description

            //In this example, we are setting the Name, DisplayName and Description to the Project namespace
            //NOTE: "Name" cannot contain spaces
            serviceBroker.Service.Name = this.GetType().Namespace;

            string friendlyName = "Google Drive File Service";
            //display name is shown to the user, so let's use a unique but friendly name
            serviceBroker.Service.MetaData.DisplayName = friendlyName;
            serviceBroker.Service.MetaData.Description = friendlyName;
        }
        #endregion

        #region void DescribeSchema()
        /// <summary>
        /// Describes the schema of the underlying data and services to the K2 platform.
        /// </summary>
        public void DescribeSchema()
        {
            //TODO: Since this is a static broker, you would add static service objects using attribute decoration.
            //The recommended approach is to create separate classes for each of your static service objects

            //in the sample implementation, we iterate over each of the classes in the assembly and if they are decorated with the 
            //ServiceObjectAttribute, we add them as service objects. 
            //if you prefer, you can manually add Service Objects like this instead: 
            //serviceBroker.Service.ServiceObjects.Add(new ServiceObject(typeof(StaticServiceObject1)));

            Type[] types = this.GetType().Assembly.GetTypes();

            foreach (Type t in types)
            {
                if (t.IsClass)
                {
                    if (t.GetCustomAttributes(typeof(SourceCode.SmartObjects.Services.ServiceSDK.Attributes.ServiceObjectAttribute), false).Length > 0)
                    {
                        this.serviceBroker.Service.ServiceObjects.Add(new SourceCode.SmartObjects.Services.ServiceSDK.Objects.ServiceObject(t));
                    }
                }
            }
        }
        #endregion

        #region void Dispose()
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            //TODO: Add any additional IDisposable implementation code here. Make sure to dispose of any data connections.

            // Clear reference to serviceBroker.
            serviceBroker = null;
        }
        #endregion

        #endregion
    }
}
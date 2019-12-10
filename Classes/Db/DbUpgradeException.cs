using System;

namespace ASTA.Classes
{
    /// <summary>
    /// Used to notify about upgrader errors.
    /// </summary>
    public class DbUpgradeException : Exception
    {
        #region Constructors

        private DbUpgradeException(DbUpgradeError error, params object[] args)
        {
            _error = error;
            _args = args;
        }

        #endregion Constructors

        #region Static Public Methods

        /// <summary>
        /// Create an unrecognized schema exception
        /// </summary>
        /// <param name="dbpath">The path to the DB file.</param>
        /// <returns></returns>
        public static DbUpgradeException UnrecognizedSchema(string dbpath)
        {
            return new DbUpgradeException(DbUpgradeError.UnrecognizedSchema, dbpath);
        }

        /// <summary>
        /// Creates an invalid upgrader schema exception
        /// </summary>
        /// <returns></returns>
        public static DbUpgradeException InvalidUpgraderSchema()
        {
            return new DbUpgradeException(DbUpgradeError.UpgraderObjectSchemasAreInvalid);
        }

        /// <summary>
        /// Creates an Internal Software Error
        /// </summary>
        /// <returns></returns>
        public static DbUpgradeException InternalSoftwareError()
        {
            return new DbUpgradeException(DbUpgradeError.InternalSoftwareError);
        }

        /// <summary>
        /// Creates a SCHEMA NOT SUPPORTED error to indicate that the
        /// schema contains DB structures that the upgrader does not have
        /// support for them.
        /// </summary>
        /// <returns></returns>
        public static DbUpgradeException SchemaIsNotSupported()
        {
            return new DbUpgradeException(DbUpgradeError.SchemaIsNotSupported);
        }

        /// <summary>
        /// Creates an invalid-primary-key-section exception.
        /// </summary>
        /// <param name="tableName">The name of the table where the section was found.</param>
        /// <returns>The exception object</returns>
        public static DbUpgradeException InvalidPrimaryKeySection(string tableName)
        {
            return new DbUpgradeException(DbUpgradeError.InvalidPrimaryKeySection, tableName);
        }

        /// <summary>
        /// Creates an upgrader module failure exception
        /// </summary>
        /// <param name="upgraderName">The name of the faulty upgrader.</param>
        /// <returns>The exception object</returns>
        public static DbUpgradeException UpgraderModuleFailure(string upgraderName)
        {
            return new DbUpgradeException(DbUpgradeError.UpgraderModuleFailure, upgraderName);
        }

        /// <summary>
        /// Creates a 'user-cancelled' exception.
        /// </summary>
        /// <returns>The exception object</returns>
        public static DbUpgradeException Cancelled()
        {
            return new DbUpgradeException(DbUpgradeError.Cancelled);
        }

        /// <summary>
        /// Creates a 'no-upgrade-path-found' exception.
        /// </summary>
        /// <param name="dbPath">The path to the DB file.</param>
        /// <returns>The exception object</returns>
        public static DbUpgradeException NoUpgradePath(string dbPath)
        {
            return new DbUpgradeException(DbUpgradeError.NoUpgradePath, dbPath);
        }

        /// <summary>
        /// Creates an exception for the case when we can't auto-upgrade a table because not all of its newer
        /// columns have a default value.
        /// </summary>
        /// <param name="tableName">The name of the problematic table</param>
        /// <returns>The exception</returns>
        public static DbUpgradeException CantAutoUpgradeATableWhenNotAllNewColumnsHaveDefaultValue(string tableName)
        {
            return new DbUpgradeException(DbUpgradeError.CantAutoUpgradeATableWhenNotAllNewColumnsHaveDefaultValue, tableName);
        }

        #endregion Static Public Methods

        #region Public Properties

        /// <summary>
        /// Return the error enumeration value
        /// </summary>
        public DbUpgradeError Error
        {
            get { return _error; }
        }

        /// <summary>
        /// Return any associated arguments
        /// </summary>
        public object[] Arguments
        {
            get { return _args; }
        }

        #endregion Public Properties

        #region Private Variabless

        private DbUpgradeError _error;
        private object[] _args;

        #endregion Private Variabless
    }

    /// <summary>
    /// Contains the possible upgrader errors.
    /// </summary>
    public enum DbUpgradeError
    {
        /// <summary>
        /// No error
        /// </summary>
        None = 0,

        /// <summary>
        /// Occurs when the upgrader object FROM schema is the same as
        /// the upgrader TO schema.
        /// </summary>
        UpgraderObjectSchemasAreInvalid = 1,

        /// <summary>
        /// Occurs when the upgrader can't recognize the schema of the
        /// DB file that we want to upgrader.
        /// </summary>
        UnrecognizedSchema = 2,

        /// <summary>
        /// Indicates an internal software error
        /// </summary>
        InternalSoftwareError = 3,

        /// <summary>
        /// DB schema is not supported in the upgrader
        /// </summary>
        SchemaIsNotSupported = 4,

        /// <summary>
        /// Invalid primary key section found in schema.
        /// </summary>
        InvalidPrimaryKeySection = 5,

        /// <summary>
        /// User cancelled the operation.
        /// </summary>
        Cancelled = 6,

        /// <summary>
        /// Failure in the upgrader module.
        /// </summary>
        UpgraderModuleFailure = 7,

        /// <summary>
        /// No upgrade path found
        /// </summary>
        NoUpgradePath = 8,

        /// <summary>
        /// Can't automatically upgrade a table because not all of its newer columns have default values.
        /// </summary>
        CantAutoUpgradeATableWhenNotAllNewColumnsHaveDefaultValue = 9,
    }
}
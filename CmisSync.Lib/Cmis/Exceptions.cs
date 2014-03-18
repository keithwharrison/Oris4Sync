using System;
using System.Runtime.Serialization;

namespace CmisSync.Lib.Cmis
{
    /// <summary>
    /// Exception launched when the CMIS server errors.
    /// </summary>
    [Serializable]
    public class BaseException : Exception
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public BaseException() { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public BaseException(string message) : base(message) { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public BaseException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public BaseException(Exception inner) : base(inner.Message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected BaseException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Exception launched when the CMIS repository denies an action.
    /// </summary>
    [Serializable]
    public class PermissionDeniedException : BaseException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public PermissionDeniedException() { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public PermissionDeniedException(string message) : base(message) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public PermissionDeniedException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public PermissionDeniedException(Exception inner) : base(inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected PermissionDeniedException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Exception launched when the CMIS server can not be found.
    /// </summary>
    [Serializable]
    public class ServerNotFoundException : BaseException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerNotFoundException() { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerNotFoundException(string message) : base(message) { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerNotFoundException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerNotFoundException(Exception inner) : base(inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected ServerNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Exception launched when user account is locked.
    /// </summary>
    [Serializable]
    public class AccountLockedException : PermissionDeniedException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public AccountLockedException() { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public AccountLockedException(string message) : base(message) { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public AccountLockedException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public AccountLockedException(Exception inner) : base(inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected AccountLockedException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Exception launched when an external user tries to access sync.
    /// </summary>
    [Serializable]
    public class ExternalUserException : PermissionDeniedException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public ExternalUserException() { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ExternalUserException(string message) : base(message) { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ExternalUserException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public ExternalUserException(Exception inner) : base(inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected ExternalUserException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Exception launched server is busy.
    /// </summary>
    [Serializable]
    public class ServerBusyException : BaseException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerBusyException() { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerBusyException(string message) : base(message) { }


        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerBusyException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        public ServerBusyException(Exception inner) : base(inner) { }

        /// <summary>
        /// Constructor.
        /// </summary>
        protected ServerBusyException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}

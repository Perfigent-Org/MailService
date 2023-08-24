using System;

namespace MailService.OAuthOffice365
{
    public class ErrorHandler
    {
        public string Message { get; set; }
    }

    public class WithExceptionOrError<T>
    {
        public WithExceptionOrError(T value, Exception ex)
        {
            Value = value;
            Exception = ex;
            ErrorMessage = ex.Message;
        }

        public WithExceptionOrError(T value, string message)
        {
            Value = value;
            ErrorMessage = message;
        }

        public WithExceptionOrError(T value) : this(value, (string)null)
        {
            Value = value;
        }

        public T Value;
        public Exception Exception;
        public string ErrorMessage;
    }
}

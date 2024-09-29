namespace EWS.Includes
{
    public static class Enumeration
    {
        public enum FlagEnum
        {
            MarkRead = 1,
            KeepUnRead = 0,
            IsAttachment = 2
        }
        public enum ResponseEnum
        {
            IntializationError = -9,
            ConfigurationError = -8,
            ExternalCredentialsError = -7,
            FormatError = -6,
            AESConfigurationError = -5,
            NotAllowedURI = -4,
            FileSystemLogFailure = -3,
            UnhandledServing = -2,
            InvalidCredentials = -1,
            Success = 0,
            UnexpectedFailure = 1
        }
    }
}
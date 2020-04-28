using System;
using System.IO;

namespace WordExport
{
    public static class Globals
    {

        public enum UploadType { TESTMANUAL, STEP }
        public enum UploadStepType { SIMPLE, VALIDATION, CALL }
        public enum UploadTestType { Acceptance, UI, E2E }
        public enum UploadPhase { NEW }  

        public const string LogPath = "logs.txt";
        public const string ErrorPath = "error.txt";

        public static void Log(string s, bool onlyToFile = false)
        {
            if (!onlyToFile)
            {
                Console.WriteLine(s);
            }
            File.AppendAllText(LogPath, s+"\r\n");
        }

        public static void Error(string s)
        {
            Console.WriteLine(s);
            File.AppendAllText(ErrorPath, s + "\r\n");
        }

        public static string AsString(this UploadType type)
        {
            switch (type)
            {
                case UploadType.TESTMANUAL:
                    return "test_manual";
                case UploadType.STEP:
                    return "step";
                default:
                    return null;
            }
        }

        public static string AsString(this UploadStepType type)
        {
            switch (type)
            {
                case UploadStepType.CALL:
                    return "Call";
                case UploadStepType.SIMPLE:
                    return "simple";
                case UploadStepType.VALIDATION:
                    return "Validation";
                default:
                    return null;
            }
        }

        public static string AsString(this UploadTestType type)
        {
            switch (type)
            {
                case UploadTestType.Acceptance:
                    return "Acceptance";
                case UploadTestType.E2E:
                    return "End to End";
                case UploadTestType.UI:
                    return "UI";
                default:
                    return null;
            }
        }

        public static string AsString(this UploadPhase phase)
        {
            switch (phase)
            {
                case UploadPhase.NEW:
                    return "New";
                default:
                    return null;
            }
        }
    }

    
}

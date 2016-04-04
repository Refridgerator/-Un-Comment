// Guids.cs
// MUST match guids.h
using System;

namespace MenuCommandUncomment
{
    static class GuidList
    {
        public const string guidMenuCommandUncommentPkgString = "b73517ac-11f9-4732-9ca2-a721869354ac";
        public const string guidMenuCommandUncommentCmdSetString = "9268c9f7-2c9d-4ac1-ac4d-85da6a52645a";

        public static readonly Guid guidMenuCommandUncommentCmdSet = new Guid(guidMenuCommandUncommentCmdSetString);
    };
}
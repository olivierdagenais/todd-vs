// Guids.cs
// MUST match guids.h
using System;

namespace SoftwareNinjas.TestOriented.VisualStudio
{
    static class GuidList
    {
        public const string PkgString = "6cc6dd01-abf1-4661-8de7-858e683cd4bc";
        public const string CmdSetString = "d87eefca-29c1-471f-bc04-0ece84ca6b82";

        public static readonly Guid CmdSet = new Guid(CmdSetString);
    };
}
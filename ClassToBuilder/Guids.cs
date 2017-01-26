// Guids.cs
// MUST match guids.h
using System;

namespace BJSS.ClassToBuilder
{
    static class GuidList
    {
        public const string guidClassToBuilderPkgString = "fcc815ed-109c-49a4-b05a-96bffc89da92";
        public const string guidClassToBuilderCmdSetString = "2a29b920-3a9f-4660-a12d-a86773854a95";

        public static readonly Guid guidClassToBuilderCmdSet = new Guid(guidClassToBuilderCmdSetString);
    };
}